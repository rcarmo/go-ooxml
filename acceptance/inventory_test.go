package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"slices"
	"strings"
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
	for _, path := range []string{overlapFeaturePath(), negativeBudgetFeaturePath(), descriptorIntegrityFeaturePath(), ownedChainFeaturePath(), crossSheetCacheFeaturePath(), runEffectsFeaturePath()} {
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
const ownedChainCaseID = "@id-xlsx-owned-calculation-chain-invalidation"
const retiredChainCaseID = "@CHAIN-001"
const crossSheetCacheCaseID = "@id-xlsx-cross-sheet-cache-invalidation"
const runEffectsCaseID = "@id-docx-go-run-effects-getters"
const runUnderlineCaseID = "@id-docx-go-run-underline-style"
const runFontNameCaseID = "@id-docx-go-run-font-name"
const retiredCacheCaseID = "@CACHE-001"

func canonicalID(path string) string {
	switch path {
	case overlapFeaturePath():
		return overlapCaseID
	case negativeBudgetFeaturePath():
		return negativeBudgetCaseID
	case descriptorIntegrityFeaturePath():
		return descriptorCollisionCaseID
	case ownedChainFeaturePath():
		return ownedChainCaseID
	case crossSheetCacheFeaturePath():
		return crossSheetCacheCaseID
	case runEffectsFeaturePath():
		return runEffectsCaseID
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
				if !canonical || (path != crossSheetCacheFeaturePath() && path != runEffectsFeaturePath()) {
					return fmt.Errorf("%s: Rules require explicit inventory support", path)
				}
				for _, child := range c.Rule.Children {
					if child.Background != nil {
						return fmt.Errorf("%s: rule backgrounds require explicit inventory support", path)
					}
					if child.Scenario != nil {
						if err := collect([]*messages.FeatureChild{{Scenario: child.Scenario}}); err != nil {
							return err
						}
					}
				}
				continue
			}
			if c.Background != nil {
				return fmt.Errorf("%s: backgrounds require explicit inventory support", path)
			}
			s := c.Scenario
			if s == nil {
				continue
			}
			id := ""
			profile := ""
			for _, tag := range s.Tags {
				if nativeIDPattern.MatchString(tag.Name) || (canonical && (tag.Name == canonicalID(path) || (path == runEffectsFeaturePath() && (tag.Name == runUnderlineCaseID || tag.Name == runFontNameCaseID)))) {
					if id != "" {
						return fmt.Errorf("%s: multiple IDs", path)
					}
					id = tag.Name
				} else if canonical && path == overlapFeaturePath() && tag.Name == "@profile-physical-member-extents" && profile == "" {
					profile = tag.Name
				} else if canonical && path == runEffectsFeaturePath() && (tag.Name == "@profile-in-memory-effects-api" || tag.Name == "@profile-document-value-api") && profile == "" {
					profile = tag.Name
				} else if !canonical || id == canonicalID(path) || (path == runEffectsFeaturePath() && (id == runUnderlineCaseID || id == runFontNameCaseID)) {
					return fmt.Errorf("%s: unsupported scenario tag %s", path, tag.Name)
				}
			}
			if canonical && id != canonicalID(path) && !(path == runEffectsFeaturePath() && (id == runUnderlineCaseID || id == runFontNameCaseID)) {
				continue // Other shared workflows remain planned for Go.
			}
			if id == "" || (canonical && path == overlapFeaturePath() && profile != "@profile-physical-member-extents") || (canonical && path == runEffectsFeaturePath() && ((id == runEffectsCaseID && profile != "@profile-in-memory-effects-api") || (id != runEffectsCaseID && profile != "@profile-document-value-api"))) {
				return fmt.Errorf("%s: scenario missing or mismatched ID/profile", path)
			}
			if !canonical && id == retiredChainCaseID {
				// Keep the historical source in the inventory, but select its canonical replacement.
				if path != filepath.Join(goFeatureRoot(), "implemented", "spreadsheet", "calc-chain.feature") || s.Name != "Dependent input edit atomically removes an owned calculation chain" || len(s.Steps) != 6 || len(s.Examples) != 0 {
					return fmt.Errorf("%s: retired CHAIN-001 source drift", path)
				}
			}
			if !canonical && id == retiredCacheCaseID {
				if path != filepath.Join(goFeatureRoot(), "implemented", "spreadsheet", "cache-invalidation.feature") || s.Name != "A cross-sheet chain invalidates transitively" || len(s.Steps) != 5 || len(s.Examples) != 0 {
					return fmt.Errorf("%s: retired CACHE-001 source drift", path)
				}
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
	if canonical && path == runEffectsFeaturePath() && (seen[runUnderlineCaseID] == "" || seen[runFontNameCaseID] == "") {
		return fmt.Errorf("%s: selected run-formatting outline missing", path)
	}
	canonicalCases := 0
	underlineRows := map[string]bool{}
	fontNameRows := map[string]bool{}
	budgetRows := map[string]bool{}
	for _, p := range gherkin.Pickles(*doc, path, next) {
		if canonical && (len(p.AstNodeIds) == 0 || (ids[p.AstNodeIds[0]] != canonicalID(path) && !(path == runEffectsFeaturePath() && (ids[p.AstNodeIds[0]] == runUnderlineCaseID || ids[p.AstNodeIds[0]] == runFontNameCaseID)))) {
			continue
		}
		line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
		id := ids[p.AstNodeIds[0]]
		if canonical {
			canonicalCases++
			if id == descriptorCollisionCaseID && (p.Name != "Unsigned descriptor geometry cannot excuse a corrupted payload" || len(p.Steps) != 8) {
				return fmt.Errorf("%s: unexpected descriptor-integrity case %q", path, p.Name)
			}
			if id == ownedChainCaseID && (p.Name != "Changing a precedent removes its owned nonstandard chain and invalidates dependent caches" || len(p.Steps) != 14) {
				return fmt.Errorf("%s: unexpected owned-chain case %q", path, p.Name)
			}
			if id == crossSheetCacheCaseID && (p.Name != "An input edit invalidates a cached answer on another sheet" || len(p.Steps) != 14) {
				return fmt.Errorf("%s: unexpected cross-sheet cache case %q", path, p.Name)
			}
			if id == runEffectsCaseID && (p.Name != "Eight direct run effects read true after setting them" || len(p.Steps) != 3) {
				return fmt.Errorf("%s: unexpected run-effects case %q", path, p.Name)
			}
			if id == runUnderlineCaseID {
				style := strings.TrimSuffix(strings.TrimPrefix(p.Name, "A "), " underline is reflected by direct getters")
				if p.Name != "A "+style+" underline is reflected by direct getters" || !slices.Contains([]string{"single", "double", "thick", "dotted", "dash", "wave"}, style) || underlineRows[style] || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word run" || p.Steps[1].Text != "its underline style is set to "+style || p.Steps[2].Text != "Underline is true and UnderlineStyle equals "+style {
					return fmt.Errorf("%s: unexpected underline row %q", path, p.Name)
				}
				underlineRows[style] = true
			}
			if id == runFontNameCaseID {
				font := strings.TrimSuffix(strings.TrimPrefix(p.Name, "A run retains direct font name "), " in memory")
				if p.Name != "A run retains direct font name "+font+" in memory" || !slices.Contains([]string{"Arial", "Times New Roman", "Calibri", "Courier New", "Georgia", "Verdana"}, font) || fontNameRows[font] || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word run" || p.Steps[1].Text != "its font name is set to "+font || p.Steps[2].Text != "its font-name getter equals "+font {
					return fmt.Errorf("%s: unexpected font-name row %q", path, p.Name)
				}
				fontNameRows[font] = true
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
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "canonical"})
		} else if id == retiredChainCaseID {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner, "selection": "superseded by " + ownedChainCaseID})
		} else if id == retiredCacheCaseID {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner, "selection": "superseded by " + crossSheetCacheCaseID})
		} else {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner})
		}
		if (lifecycle == "@implemented" && id != retiredChainCaseID && id != retiredCacheCaseID) || canonical {
			expected[key] = expectedCase{key, p.Name, len(p.Steps)}
		}
	}
	if canonical && ((canonicalID(path) == negativeBudgetCaseID && (canonicalCases != 2 || len(budgetRows) != 2)) || (canonicalID(path) == overlapCaseID && canonicalCases != 1) || (canonicalID(path) == descriptorCollisionCaseID && canonicalCases != 1) || (canonicalID(path) == ownedChainCaseID && canonicalCases != 1) || (canonicalID(path) == crossSheetCacheCaseID && canonicalCases != 1) || (canonicalID(path) == runEffectsCaseID && (canonicalCases != 13 || len(underlineRows) != 6 || len(fontNameRows) != 6))) {
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
