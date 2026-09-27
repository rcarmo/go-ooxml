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
	err := filepath.WalkDir(goFeatureRoot(), func(path string, d os.DirEntry, err error) error {
		return inventoryFeature(path, d, err, expected, &inventory, seen, next, false)
	})
	if err != nil {
		return nil, nil, err
	}
	for _, path := range []string{overlapFeaturePath(), negativeBudgetFeaturePath(), descriptorIntegrityFeaturePath()} {
		if err = inventoryFeature(path, nil, nil, expected, &inventory, seen, next, true); err != nil {
			return nil, nil, err
		}
	}
	return expected, inventory, nil
}

var nativeIDPattern = regexp.MustCompile(`^@[A-Z]+-[0-9]{3}$`)

const overlapCaseID = "@id-zip-physical-member-overlap-refusal"
const negativeBudgetCaseID = "@id-package-admission-negative-budget"
const descriptorCollisionCaseID = "@id-zip-unsigned-descriptor-signature-collision"

func canonicalID(path string) string {
	switch path {
	case overlapFeaturePath():
		return overlapCaseID
	case negativeBudgetFeaturePath():
		return negativeBudgetCaseID
	case descriptorIntegrityFeaturePath():
		return descriptorCollisionCaseID
	default:
		return ""
	}
}

func inventoryFeature(path string, d os.DirEntry, err error, expected map[caseID]expectedCase, inventory *[]map[string]any, seen map[string]string, next func() string, canonical bool) error {
	if err != nil {
		return err
	}
	if (d != nil && d.IsDir()) || filepath.Ext(path) != ".feature" {
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
	if canonical {
		if canonicalID(path) == "" || lifecycle != "@planned" || runner != "" {
			return fmt.Errorf("%s: unexpected canonical feature metadata", path)
		}
	} else if lifecycle == "" || runner == "" {
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
			profile := false
			for _, tag := range s.Tags {
				if nativeIDPattern.MatchString(tag.Name) || (canonical && tag.Name == canonicalID(path)) {
					if id != "" {
						return fmt.Errorf("%s: multiple IDs", path)
					}
					id = tag.Name
				} else if canonical && path == overlapFeaturePath() && tag.Name == "@profile-physical-member-extents" && !profile {
					profile = true
				} else if !canonical || id == canonicalID(path) {
					return fmt.Errorf("%s: unsupported scenario tag %s", path, tag.Name)
				}
			}
			if canonical && id != canonicalID(path) {
				continue // Other shared workflows remain planned for Go.
			}
			if id == "" || (canonical && path == overlapFeaturePath() && !profile) {
				return fmt.Errorf("%s: scenario missing ID/profile", path)
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
	if canonical && seen[canonicalID(path)] == "" {
		return fmt.Errorf("%s: selected canonical case missing", path)
	}
	canonicalCases := 0
	budgetRows := map[string]bool{}
	for _, p := range gherkin.Pickles(*doc, path, next) {
		if canonical && (len(p.AstNodeIds) == 0 || ids[p.AstNodeIds[0]] != canonicalID(path)) {
			continue
		}
		line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
		id := ids[p.AstNodeIds[0]]
		if canonical {
			canonicalCases++
			if id == descriptorCollisionCaseID && (p.Name != "Unsigned descriptor geometry cannot excuse a corrupted payload" || len(p.Steps) != 8) {
				return fmt.Errorf("%s: unexpected descriptor-integrity case %q", path, p.Name)
			}
			if id == negativeBudgetCaseID {
				if (p.Name != "A negative source bytes budget refuses before package intake" && p.Name != "A negative entry count budget refuses before package intake") || budgetRows[p.Name] || len(p.Steps) != 6 {
					return fmt.Errorf("%s: unexpected negative-budget case %q", path, p.Name)
				}
				budgetRows[p.Name] = true
			}
		}
		key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
		if canonical {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go"})
		} else {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner})
		}
		if lifecycle == "@implemented" || canonical {
			expected[key] = expectedCase{key, p.Name, len(p.Steps)}
		}
	}
	if canonical && ((canonicalID(path) == negativeBudgetCaseID && (canonicalCases != 2 || len(budgetRows) != 2)) || (canonicalID(path) == overlapCaseID && canonicalCases != 1) || (canonicalID(path) == descriptorCollisionCaseID && canonicalCases != 1)) {
		return fmt.Errorf("%s: selected canonical case count drift: %d", path, canonicalCases)
	}
	return nil
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
