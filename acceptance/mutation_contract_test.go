package acceptance

import (
	"bytes"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/testutil"
)

type mutationFixture struct {
	ID      string            `json:"id"`
	AssetID string            `json:"assetId"`
	Facts   map[string]any    `json:"facts"`
	Allowed []string          `json:"allowedChangedPartsForSuccess"`
	Members map[string]string `json:"memberSha256"`
}
type mutationContract struct {
	Schema    int      `json:"schemaVersion"`
	Revision  string   `json:"contractRevision"`
	Feature   string   `json:"feature"`  // schema 1 only
	Features  []string `json:"features"` // schema 2 only
	Scenarios []string `json:"scenarioIds"`
	Count     int      `json:"expandedCaseCount"`
	Policy    struct {
		Membership string `json:"membership"`
		Preserve   string `json:"preserve"`
	} `json:"fixturePolicy"`
	Fixtures []mutationFixture `json:"fixtures"`
}
type mutationStep struct {
	Text     string          `json:"text"`
	Argument json.RawMessage `json:"argument"`
}
type mutationCase struct {
	ID       string            `json:"stableCaseKey"`
	Scenario string            `json:"scenarioId"`
	Examples map[string]string `json:"examples"`
	Steps    []mutationStep    `json:"expandedSteps"`
}

func parseMutationContract(b []byte) (mutationContract, error) {
	var c mutationContract
	d := json.NewDecoder(bytes.NewReader(b))
	d.DisallowUnknownFields()
	if err := d.Decode(&c); err != nil {
		return c, err
	}
	if err := d.Decode(new(any)); err != io.EOF {
		return c, fmt.Errorf("trailing contract content")
	}
	if (c.Schema != 1 && c.Schema != 2) || c.Revision != "ooxml-shared-contracts-v2" || c.Count != 19 || len(c.Scenarios) != 8 || len(c.Fixtures) != 4 || c.Policy.Membership != "exact" || c.Policy.Preserve != "all-except-allowed" {
		return c, fmt.Errorf("unsupported mutation contract")
	}
	if c.Schema == 1 {
		if c.Feature != "workflows/mutation-safety.feature" || len(c.Features) != 0 {
			return c, fmt.Errorf("invalid schema 1 mutation feature")
		}
	} else {
		if c.Feature != "" || len(c.Features) == 0 {
			return c, fmt.Errorf("schema 2 requires an explicit feature list")
		}
		paths := map[string]bool{}
		for _, feature := range c.Features {
			if !strings.HasPrefix(feature, "workflows/") || !strings.HasSuffix(feature, ".feature") || !filepath.IsLocal(feature) || path.Clean(feature) != feature || strings.Contains(feature, "\\") || paths[feature] {
				return c, fmt.Errorf("invalid/duplicate mutation feature %s", feature)
			}
			paths[feature] = true
		}
	}
	seen := map[string]bool{}
	assets := map[string]bool{}
	for _, id := range c.Scenarios {
		if !strings.HasPrefix(id, "@id-") || seen[id] {
			return c, fmt.Errorf("invalid/duplicate scenario %s", id)
		}
		seen[id] = true
	}
	seen = map[string]bool{}
	for _, f := range c.Fixtures {
		if f.ID == "" || seen[f.ID] || assets[f.AssetID] || !strings.HasPrefix(f.AssetID, "fixture-") || len(f.AssetID) != 72 || len(f.Facts) == 0 || len(f.Members) == 0 {
			return c, fmt.Errorf("invalid mutation fixture %s", f.ID)
		}
		if _, err := hex.DecodeString(strings.TrimPrefix(f.AssetID, "fixture-")); err != nil {
			return c, err
		}
		seen[f.ID] = true
		assets[f.AssetID] = true
		for name, hash := range f.Members {
			if name == "" || len(hash) != 64 || path.Clean(name) != name || strings.HasPrefix(name, "/") || strings.Contains(name, "\\") || name == ".." || strings.HasPrefix(name, "../") {
				return c, fmt.Errorf("invalid member hash %s", name)
			}
			if _, err := hex.DecodeString(hash); err != nil {
				return c, err
			}
		}
		if _, err := preservedMembers(f); err != nil {
			return c, err
		}
	}
	return c, nil
}
func preservedMembers(f mutationFixture) (map[string]string, error) {
	allowed := map[string]bool{}
	for _, name := range f.Allowed {
		if _, ok := f.Members[name]; !ok || allowed[name] {
			return nil, fmt.Errorf("unknown/duplicate allowed part %s", name)
		}
		allowed[name] = true
	}
	keep := map[string]string{}
	for name, hash := range f.Members {
		if !allowed[name] {
			keep[name] = hash
		}
	}
	return keep, nil
}

// Compile official Gherkin, deriving examples from AST row identities and typed
// tables from pickle arguments. No generated expanded-contracts input is read.
func compileMutationCases(b []byte, path string, ids []string, count int) ([]mutationCase, error) {
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(bytes.NewReader(b), next)
	if err != nil {
		return nil, err
	}
	if doc.Feature == nil {
		return nil, fmt.Errorf("missing feature")
	}
	expected := map[string]bool{}
	for _, id := range ids {
		if expected[id] {
			return nil, fmt.Errorf("duplicate contract ID")
		}
		expected[id] = true
	}
	values := map[string]map[string]string{}
	owners := map[string]string{}
	defined := map[string]bool{}
	indexScenario := func(s *messages.Scenario) error {
		id := ""
		for _, tag := range s.Tags {
			if strings.HasPrefix(tag.Name, "@id-") {
				if id != "" {
					return fmt.Errorf("multiple scenario IDs")
				}
				id = tag.Name
			}
		}
		if !expected[id] {
			if count != 0 { // schema 1 owned the whole feature
				return fmt.Errorf("unknown workflow %s", id)
			}
			return nil // other planned scenarios in an operation-group feature
		}
		if defined[id] {
			return fmt.Errorf("duplicate workflow %s", id)
		}
		defined[id] = true
		owners[s.Id] = id
		values[s.Id] = map[string]string{}
		for _, ex := range s.Examples {
			if ex.TableHeader == nil || len(ex.TableBody) == 0 {
				return fmt.Errorf("empty examples")
			}
			headers := map[string]bool{}
			for _, h := range ex.TableHeader.Cells {
				if h.Value == "" || headers[h.Value] {
					return fmt.Errorf("duplicate example column")
				}
				headers[h.Value] = true
			}
			for _, row := range ex.TableBody {
				if len(row.Cells) != len(ex.TableHeader.Cells) {
					return fmt.Errorf("example width mismatch")
				}
				m := map[string]string{}
				for i, cell := range row.Cells {
					m[ex.TableHeader.Cells[i].Value] = cell.Value
				}
				values[row.Id] = m
			}
		}
		return nil
	}
	for _, child := range doc.Feature.Children {
		if child.Scenario != nil {
			if err := indexScenario(child.Scenario); err != nil {
				return nil, err
			}
		} else if child.Rule != nil && count == 0 {
			for _, nested := range child.Rule.Children {
				if nested.Scenario == nil {
					return nil, fmt.Errorf("unsupported mutation rule child")
				}
				if err := indexScenario(nested.Scenario); err != nil {
					return nil, err
				}
			}
		} else {
			return nil, fmt.Errorf("unsupported shared feature child")
		}
	}
	if len(defined) != len(expected) {
		return nil, fmt.Errorf("missing declared workflow")
	}
	result := []mutationCase{}
	seen := map[string]bool{}
	for _, p := range gherkin.Pickles(*doc, path, next) {
		if len(p.AstNodeIds) == 0 {
			return nil, fmt.Errorf("missing pickle identity")
		}
		id := owners[p.AstNodeIds[0]]
		if id == "" && count == 0 {
			continue // unselected planned operation-group scenario
		}
		m, ok := values[p.AstNodeIds[len(p.AstNodeIds)-1]]
		if id == "" || !ok {
			return nil, fmt.Errorf("unknown pickle identity")
		}
		key := caseKey(id, m)
		if seen[key] {
			return nil, fmt.Errorf("duplicate case key %s", key)
		}
		seen[key] = true
		planned := false
		for _, tag := range p.Tags {
			if tag.Name == "@planned" {
				planned = true
			}
			if tag.Name == "@implemented" || tag.Name == "@external" {
				return nil, fmt.Errorf("shared inventory claims execution lifecycle")
			}
		}
		if !planned {
			return nil, fmt.Errorf("shared case not planned")
		}
		c := mutationCase{ID: key, Scenario: id, Examples: m, Steps: []mutationStep{}}
		for _, s := range p.Steps {
			arg, err := sharedArgument(s.Argument)
			if err != nil {
				return nil, err
			}
			if err = validateTypedTable(arg); err != nil {
				return nil, err
			}
			c.Steps = append(c.Steps, mutationStep{Text: s.Text, Argument: arg})
		}
		result = append(result, c)
	}
	if count != 0 && len(result) != count {
		return nil, fmt.Errorf("expanded cases %d != %d", len(result), count)
	}
	return result, nil
}
func sharedArgument(a *messages.PickleStepArgument) (json.RawMessage, error) {
	if a == nil {
		return nil, nil
	}
	if a.DataTable != nil {
		rows := [][]string{}
		for _, r := range a.DataTable.Rows {
			row := []string{}
			for _, c := range r.Cells {
				row = append(row, c.Value)
			}
			rows = append(rows, row)
		}
		return json.Marshal(map[string]any{"dataTable": rows})
	}
	if a.DocString != nil {
		return json.Marshal(map[string]any{"docString": map[string]string{"content": a.DocString.Content, "mediaType": a.DocString.MediaType}})
	}
	return nil, fmt.Errorf("unrecognised pickle argument")
}

func loadMutationContract(pin referencePin) (mutationContract, []mutationCase, error) {
	// The released v0.2 adapter remains until the next tag is verified. Selection
	// is pinned schema, never filesystem presence or a silent fallback.
	if pin.Schema == 1 {
		return loadLegacyMutation(pin)
	}
	if pin.Schema != 2 {
		return mutationContract{}, nil, fmt.Errorf("unknown reference schema")
	}
	b, err := os.ReadFile(testutil.ReferencePath("contracts", "mutation-safety.json"))
	if err != nil {
		return mutationContract{}, nil, err
	}
	c, err := parseMutationContract(b)
	if err != nil {
		return c, nil, err
	}
	if c.Schema == 1 {
		feature, err := os.ReadFile(testutil.ReferencePath(c.Feature))
		if err != nil {
			return c, nil, err
		}
		cases, err := compileMutationCases(feature, c.Feature, c.Scenarios, c.Count)
		return c, cases, err
	}
	cases, err := loadSplitMutation(c)
	return c, cases, err
}

// Compile the selected workflows from each declared feature. Other planned
// scenarios in the same operation feature do not become mutation contracts.
func loadSplitMutation(c mutationContract) ([]mutationCase, error) {
	return compileSplitMutation(c, func(path string) ([]byte, error) {
		return os.ReadFile(testutil.ReferencePath(path))
	})
}

func compileSplitMutation(c mutationContract, read func(string) ([]byte, error)) ([]mutationCase, error) {
	declared := make(map[string]bool, len(c.Scenarios))
	for _, id := range c.Scenarios {
		declared[id] = true
	}
	owners := map[string]string{}
	var cases []mutationCase
	seen := map[string]bool{}
	for _, featurePath := range c.Features {
		b, err := read(featurePath)
		if err != nil {
			return nil, err
		}
		ids, err := mutationFeatureIDs(b, featurePath, declared, owners)
		if err != nil {
			return nil, err
		}
		if len(ids) == 0 {
			return nil, fmt.Errorf("mutation feature has no selected scenarios: %s", featurePath)
		}
		compiled, err := compileMutationCases(b, featurePath, ids, 0)
		if err != nil {
			return nil, err
		}
		for _, row := range compiled {
			if seen[row.ID] {
				return nil, fmt.Errorf("duplicate mutation case key %s", row.ID)
			}
			seen[row.ID] = true
			cases = append(cases, row)
		}
	}
	if len(owners) != len(declared) || len(cases) != c.Count {
		return nil, fmt.Errorf("mutation inventory differs: %d cases, %d of %d IDs", len(cases), len(owners), len(declared))
	}
	return cases, nil
}

func mutationFeatureIDs(b []byte, featurePath string, declared map[string]bool, owners map[string]string) ([]string, error) {
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(bytes.NewReader(b), next)
	if err != nil {
		return nil, err
	}
	if doc.Feature == nil {
		return nil, fmt.Errorf("missing mutation feature %s", featurePath)
	}
	var ids []string
	for _, child := range doc.Feature.Children {
		if child.Scenario != nil {
			id := ""
			for _, tag := range child.Scenario.Tags {
				if strings.HasPrefix(tag.Name, "@id-") {
					if id != "" {
						return nil, fmt.Errorf("multiple scenario IDs in %s", featurePath)
					}
					id = tag.Name
				}
			}
			if declared[id] {
				if owners[id] != "" {
					return nil, fmt.Errorf("mutation scenario %s appears in both %s and %s", id, owners[id], featurePath)
				}
				owners[id] = featurePath
				ids = append(ids, id)
			}
		} else if child.Rule != nil {
			for _, nested := range child.Rule.Children {
				if nested.Scenario == nil {
					continue
				}
				id := ""
				for _, tag := range nested.Scenario.Tags {
					if strings.HasPrefix(tag.Name, "@id-") {
						if id != "" {
							return nil, fmt.Errorf("multiple scenario IDs in %s", featurePath)
						}
						id = tag.Name
					}
				}
				if declared[id] {
					if owners[id] != "" {
						return nil, fmt.Errorf("mutation scenario %s appears in both %s and %s", id, owners[id], featurePath)
					}
					owners[id] = featurePath
					ids = append(ids, id)
				}
			}
		}
	}
	return ids, nil
}
func loadLegacyMutation(pin referencePin) (mutationContract, []mutationCase, error) {
	root := testutil.ReferencePath("shared", "v2", "pack")
	b, err := pinnedFile(root, "pack-manifest.json", pin.Pack)
	if err != nil {
		return mutationContract{}, nil, err
	}
	var pack struct {
		Revision string            `json:"contractRevision"`
		Files    map[string]string `json:"files"`
	}
	if err = json.Unmarshal(b, &pack); err != nil {
		return mutationContract{}, nil, err
	}
	for name, hash := range pack.Files {
		if _, err = pinnedFile(root, name, hash); err != nil {
			return mutationContract{}, nil, err
		}
	}
	b, err = os.ReadFile(testutil.ReferencePath("shared", "v2", "pack", "fixture-manifest.json"))
	if err != nil {
		return mutationContract{}, nil, err
	}
	var fm struct {
		Fixtures []struct {
			ID      string            `json:"id"`
			AssetID string            `json:"assetId"`
			Facts   map[string]any    `json:"facts"`
			Allowed []string          `json:"allowedChangedPartsForSuccess"`
			Members map[string]string `json:"memberSha256"`
		} `json:"fixtures"`
	}
	if err = json.Unmarshal(b, &fm); err != nil {
		return mutationContract{}, nil, err
	}
	b, err = os.ReadFile(testutil.ReferencePath("shared", "v2", "pack", "expanded-contracts.json"))
	if err != nil {
		return mutationContract{}, nil, err
	}
	var legacy struct {
		Cases []mutationCase `json:"inventory"`
	}
	if err = json.Unmarshal(b, &legacy); err != nil {
		return mutationContract{}, nil, err
	}
	c := mutationContract{Schema: 1, Revision: pack.Revision, Feature: "workflows/mutation-safety.feature", Count: 19}
	c.Policy.Membership = "exact"
	c.Policy.Preserve = "all-except-allowed"
	seen := map[string]bool{}
	for _, row := range legacy.Cases {
		if !seen[row.Scenario] {
			seen[row.Scenario] = true
			c.Scenarios = append(c.Scenarios, row.Scenario)
		}
	}
	for _, f := range fm.Fixtures {
		c.Fixtures = append(c.Fixtures, mutationFixture{ID: f.ID, AssetID: f.AssetID, Facts: f.Facts, Allowed: f.Allowed, Members: f.Members})
	}
	check, _ := json.Marshal(c)
	c, err = parseMutationContract(check)
	if err != nil {
		return c, nil, err
	}
	feature, err := os.ReadFile(testutil.ReferencePath("shared", "v2", "pack", "features", "mutation-safety.feature"))
	if err != nil {
		return c, nil, err
	}
	cases, err := compileMutationCases(feature, c.Feature, c.Scenarios, c.Count)
	if err != nil {
		return c, nil, err
	}
	if len(legacy.Cases) != len(cases) {
		return c, nil, fmt.Errorf("legacy case count differs")
	}
	byID := map[string]mutationCase{}
	for _, r := range legacy.Cases {
		if _, exists := byID[r.ID]; exists {
			return c, nil, fmt.Errorf("duplicate legacy case identity")
		}
		byID[r.ID] = r
	}
	for _, r := range cases {
		a, ok := byID[r.ID]
		if !ok || len(a.Steps) != len(r.Steps) {
			return c, nil, fmt.Errorf("compiled legacy identity differs")
		}
		for i, s := range r.Steps {
			if s.Text != a.Steps[i].Text {
				return c, nil, fmt.Errorf("compiled legacy step differs")
			}
			var left, right any
			if len(s.Argument) != 0 {
				if err := json.Unmarshal(s.Argument, &left); err != nil {
					return c, nil, err
				}
			}
			if len(a.Steps[i].Argument) != 0 {
				if err := json.Unmarshal(a.Steps[i].Argument, &right); err != nil {
					return c, nil, err
				}
			}
			if !reflect.DeepEqual(left, right) {
				return c, nil, fmt.Errorf("compiled legacy argument differs")
			}
		}
	}
	return c, cases, nil
}

func TestMutationContractAndCompilerBatch(t *testing.T) {
	pin := loadReferencePin(t)
	c, cases, err := loadMutationContract(pin)
	if err != nil {
		t.Fatal(err)
	}
	raw, _ := json.Marshal(c)
	for _, tc := range []struct {
		name   string
		change func(*mutationContract)
	}{
		{"unknown membership", func(x *mutationContract) { x.Policy.Membership = "partial" }},
		{"duplicate scenario", func(x *mutationContract) { x.Scenarios[1] = x.Scenarios[0] }},
		{"duplicate fixture", func(x *mutationContract) { x.Fixtures[1] = x.Fixtures[0] }},
		{"unknown allowed member", func(x *mutationContract) { x.Fixtures[0].Allowed = append(x.Fixtures[0].Allowed, "missing.xml") }},
		{"duplicate allowed member", func(x *mutationContract) {
			x.Fixtures[0].Allowed = append(x.Fixtures[0].Allowed, x.Fixtures[0].Allowed[0])
		}},
		{"malformed member hash", func(x *mutationContract) { x.Fixtures[0].Members["bad.xml"] = strings.Repeat("z", 64) }},
		{"schema 1 rejects feature list", func(x *mutationContract) {
			if x.Schema == 1 {
				x.Features = []string{x.Feature}
			} else {
				x.Feature = "workflows/mutation-safety.feature"
			}
		}},
		{"schema 2 requires feature list", func(x *mutationContract) {
			if x.Schema == 2 {
				x.Features = nil
			} else {
				x.Feature = ""
			}
		}},
		{"schema 2 rejects duplicate feature", func(x *mutationContract) {
			if x.Schema == 2 {
				x.Features = append(x.Features, x.Features[0])
			} else {
				x.Features = []string{x.Feature, x.Feature}
			}
		}},
		{"schema 2 rejects escaping feature", func(x *mutationContract) {
			if x.Schema == 2 {
				x.Features[0] = "workflows/../outside.feature"
			} else {
				x.Features = []string{"workflows/../outside.feature"}
			}
		}},
		{"schema 2 rejects absolute feature", func(x *mutationContract) {
			if x.Schema == 2 {
				x.Features[0] = "/tmp/outside.feature"
			} else {
				x.Features = []string{"/tmp/outside.feature"}
			}
		}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			var x mutationContract
			if err := json.Unmarshal(raw, &x); err != nil {
				t.Fatal(err)
			}
			tc.change(&x)
			b, _ := json.Marshal(x)
			if _, err := parseMutationContract(b); err == nil {
				t.Fatal("invalid contract accepted")
			}
		})
	}
	for _, f := range c.Fixtures {
		keep, err := preservedMembers(f)
		if err != nil {
			t.Fatal(err)
		}
		if len(keep)+len(f.Allowed) != len(f.Members) {
			t.Fatal("preservation complement incomplete")
		}
		for _, name := range f.Allowed {
			if _, ok := keep[name]; ok {
				t.Fatal("allowed member treated as preserved")
			}
		}
	}
	if c.Schema == 1 {
		featurePath := testutil.ReferencePath(c.Feature)
		if pin.Schema == 1 {
			featurePath = testutil.ReferencePath("shared", "v2", "pack", "features", "mutation-safety.feature")
		}
		feature, err := os.ReadFile(featurePath)
		if err != nil {
			t.Fatal(err)
		}
		for _, tc := range []struct {
			name  string
			data  []byte
			ids   []string
			count int
		}{
			{"wrong expanded count", feature, c.Scenarios, c.Count + 1},
			{"unknown scenario", bytes.Replace(feature, []byte(c.Scenarios[0]), []byte("@id-absent"), 1), c.Scenarios, c.Count},
			{"claim executed lifecycle", bytes.Replace(feature, []byte("@planned"), []byte("@implemented"), 1), c.Scenarios, c.Count},
			{"duplicate destination row", bytes.ReplaceAll(feature, []byte("distinct-absent"), []byte("source")), c.Scenarios, c.Count},
			{"invalid typed JSON", bytes.ReplaceAll(feature, []byte(`"Changed by dry run"`), []byte(`NaN`)), c.Scenarios, c.Count},
		} {
			t.Run(tc.name, func(t *testing.T) {
				if _, err := compileMutationCases(tc.data, c.Feature, tc.ids, tc.count); err == nil {
					t.Fatal("invalid feature inventory accepted")
				}
			})
		}
	} else {
		feature := c.Features[0]
		original, err := os.ReadFile(testutil.ReferencePath(feature))
		if err != nil {
			t.Fatal(err)
		}
		fixtureBytes := map[string][]byte{}
		for _, name := range c.Features {
			fixtureBytes[name], err = os.ReadFile(testutil.ReferencePath(name))
			if err != nil {
				t.Fatal(err)
			}
		}
		for _, tc := range []struct {
			name   string
			change func(*mutationContract)
			input  []byte
		}{
			{"missing selected scenario", nil, bytes.Replace(original, []byte(c.Scenarios[1]), []byte("@id-absent"), 1)},
			{"claim executed lifecycle", nil, bytes.Replace(original, []byte("@planned"), []byte("@implemented"), 1)},
			{"duplicate selected scenario in second feature", func(x *mutationContract) { x.Features = append(x.Features, feature+"-duplicate.feature") }, original},
			{"wrong expanded count", func(x *mutationContract) { x.Count++ }, nil},
			{"duplicate destination row", nil, bytes.ReplaceAll(original, []byte("distinct-absent"), []byte("source"))},
			{"invalid typed JSON", nil, bytes.ReplaceAll(original, []byte(`"changed"`), []byte(`NaN`))},
		} {
			t.Run(tc.name, func(t *testing.T) {
				x := c
				if tc.change != nil {
					tc.change(&x)
				}
				changed := make(map[string][]byte, len(fixtureBytes)+1)
				for path, contents := range fixtureBytes {
					changed[path] = contents
				}
				if tc.input != nil {
					changed[feature] = tc.input
				}
				if strings.Contains(tc.name, "duplicate selected") {
					changed[feature+"-duplicate.feature"] = original
				}
				if _, err := compileSplitMutation(x, func(name string) ([]byte, error) { return changed[name], nil }); err == nil {
					t.Fatal("invalid split mutation inventory accepted")
				}
			})
		}
	}
	if len(cases) != 19 {
		t.Fatal("case inventory")
	}
	seen := map[string]bool{}
	for _, row := range cases {
		seen[row.ID] = true
	}
	if !seen["@id-xlsx-cross-sheet-cache-invalidation:{}"] || !seen[`@id-pptx-dry-run-no-mutation:{"destination":"source"}`] {
		t.Fatal("stable keys changed")
	}
}
