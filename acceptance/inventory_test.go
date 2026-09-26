package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
)

// Inventory all feature files, including planned and external contracts. Native
// execution selects only implemented cases; excluded cases remain in the map.
func inventoryCases() (map[caseID]expectedCase, []map[string]any, error) {
	expected := map[caseID]expectedCase{}
	var inventory []map[string]any
	seen := map[string]string{}
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	pattern := regexp.MustCompile(`^@[A-Z]+-[0-9]{3}$`)
	err := filepath.WalkDir("../features", func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() || filepath.Ext(path) != ".feature" {
			return nil
		}
		f, err := os.Open(path)
		if err != nil {
			return err
		}
		doc, err := gherkin.ParseGherkinDocument(f, next)
		_ = f.Close()
		if err != nil {
			return err
		}
		if doc.Feature == nil {
			return fmt.Errorf("%s: no feature", path)
		}
		lifecycle, runner := "", ""
		for _, tag := range doc.Feature.Tags {
			switch tag.Name {
			case "@implemented", "@planned", "@external":
				if lifecycle != "" {
					return fmt.Errorf("%s: multiple lifecycles", path)
				}
				lifecycle = tag.Name
			case "@go", "@office":
				if runner != "" {
					return fmt.Errorf("%s: multiple runners", path)
				}
				runner = tag.Name
			}
		}
		if lifecycle == "" || runner == "" {
			return fmt.Errorf("%s: lifecycle/runner missing", path)
		}
		if lifecycle == "@implemented" && runner != "@go" {
			return fmt.Errorf("%s: implemented case omitted by native runner", path)
		}
		if lifecycle == "@external" && runner != "@office" {
			return fmt.Errorf("%s: external runner mismatch", path)
		}
		lines := map[string]int{}
		ids := map[string]string{}
		var collect func([]*messages.FeatureChild) error
		collect = func(children []*messages.FeatureChild) error {
			for _, c := range children {
				if c.Rule != nil {
					return fmt.Errorf("%s: Rules require explicit inventory support", path)
				}
				if c.Background != nil {
					return fmt.Errorf("%s: backgrounds require explicit inventory support", path)
				}
				s := c.Scenario
				if s == nil {
					continue
				}
				id := ""
				for _, tag := range s.Tags {
					if pattern.MatchString(tag.Name) {
						if id != "" {
							return fmt.Errorf("%s: multiple IDs", path)
						}
						id = tag.Name
					} else {
						return fmt.Errorf("%s: unsupported scenario tag %s", path, tag.Name)
					}
				}
				if id == "" {
					return fmt.Errorf("%s: scenario missing ID", path)
				}
				if prior, ok := seen[id]; ok {
					return fmt.Errorf("duplicate ID %s in %s and %s", id, prior, path)
				}
				seen[id] = path
				lines[s.Id] = int(s.Location.Line)
				ids[s.Id] = id
				if len(s.Steps) == 0 {
					return fmt.Errorf("%s: empty scenario", path)
				}
				for _, ex := range s.Examples {
					if len(ex.Tags) != 0 {
						return fmt.Errorf("%s: Examples tags not supported yet", path)
					}
					for _, row := range ex.TableBody {
						lines[row.Id] = int(row.Location.Line)
					}
				}
			}
			return nil
		}
		if err := collect(doc.Feature.Children); err != nil {
			return err
		}
		for _, p := range gherkin.Pickles(*doc, path, next) {
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			id := ids[p.AstNodeIds[0]]
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			inventory = append(inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner})
			if lifecycle == "@implemented" {
				expected[key] = expectedCase{key, p.Name, len(p.Steps)}
			}
		}
		return nil
	})
	return expected, inventory, err
}

func TestResultReconciliation(t *testing.T) {
	key := caseID{"f.feature", "@TEST-001", 12}
	expected := map[caseID]expectedCase{key: {key, "example", 1}}
	fixture := func(status string) []byte {
		return []byte(fmt.Sprintf(`[{"uri":"f.feature","elements":[{"name":"example","line":12,"type":"scenario","tags":[{"name":"@TEST-001"}],"steps":[{"name":"observed","result":{"status":%q}}]}]}]`, status))
	}
	cases := []struct {
		name      string
		data      []byte
		wantError bool
	}{
		{"passed", fixture("passed"), false}, {"missing", []byte(`[]`), true}, {"skipped", fixture("skipped"), true}, {"undefined", fixture("undefined"), true}, {"pending", fixture("pending"), true}, {"failed", fixture("failed"), true}, {"invalid", []byte(`no`), true},
	}
	var duplicated []reportFeature
	_ = json.Unmarshal(fixture("passed"), &duplicated)
	duplicated[0].Elements = append(duplicated[0].Elements, duplicated[0].Elements[0])
	dup, _ := json.Marshal(duplicated)
	cases = append(cases, struct {
		name      string
		data      []byte
		wantError bool
	}{"duplicate", dup, true})
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			err := reconcile(expected, tc.data)
			if (err != nil) != tc.wantError {
				t.Fatalf("error=%v expected failure=%v", err, tc.wantError)
			}
		})
	}
}
