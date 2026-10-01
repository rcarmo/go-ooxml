package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// The batch-2 lane is opt-in against the one sealed checkout, never the
// published default reference distribution or adjacent planned scenarios.
const batch2Commit = "14a7bf7ad72baa41028ff140802cbbfdbbd8a845"
const batch2Manifest = "b6a3548a69bfd1524a1b80899538e25db7b4ec51431d39fc8d5f4420b9444539"

var batch2Paths = map[string][]string{
	"workflows/package/zip32.feature": {
		"@id-zip-crc32-standard-vector", "@id-zip-read-valid", "@id-zip-refuse-unsafe", "@id-zip-bounds", "@id-zip-write-deterministic",
	},
	"workflows/package/zip-admission.feature":        {"@id-package-admission-resource-limits"},
	"workflows/package/xml-member-admission.feature": {"@id-package-admission-unsafe-xml-members"},
	"workflows/package/preservation.feature": {
		"@id-opc-package-corpus-noop", "@id-bun-opc-detached-byte-copies", "@id-bun-opc-preserve-utf16le-edit", "@id-opc-package-transaction-rollback", "@id-opc-package-preserve-unrelated",
	},
	"workflows/package/graph.feature":                   {"@id-opc-add-related-part", "@id-opc-graph-rollback", "@id-opc-remove-related-part", "@id-opc-diff-content-type"},
	"workflows/package/semantic-diff.feature":           {"@id-package-diff-equivalent-xml-and-binary-changes"},
	"workflows/package/relationship-namespaces.feature": {"@id-office-relationship-prefix-alias", "@id-office-relationship-wrong-uri"},
	"workflows/xml/comparison.feature":                  {"@id-xml-comparison-prefix-and-opc-order"},
}

func batch2SelectedID(id string) bool {
	for _, ids := range batch2Paths {
		for _, selected := range ids {
			if id == selected {
				return true
			}
		}
	}
	return false
}

const batch2Root = "/workspace/projects/fixtures-ooxml"

func batch2PinMatches(pin referencePin) bool {
	return pin.Schema == 2 && pin.Commit == batch2Commit && pin.Manifest == batch2Manifest && pin.TagObject == "" && pin.Tag == "candidate-package-alignment"
}

// A clean clone at the same commit is not the reserved reviewed checkout.
// Resolve aliases, but require the exact physical candidate root.
func batch2RootAllowed(root string) error {
	want, err := filepath.EvalSymlinks(batch2Root)
	if err != nil {
		return err
	}
	got, err := filepath.EvalSymlinks(root)
	if err != nil {
		return err
	}
	want, err = filepath.Abs(want)
	if err != nil {
		return err
	}
	got, err = filepath.Abs(got)
	if err != nil {
		return err
	}
	if got != want {
		return fmt.Errorf("batch-2 candidate root %s differs from reviewed root %s", got, want)
	}
	return nil
}

func batch2Candidate() bool {
	path := os.Getenv("OOXML_REFERENCE_PIN")
	if path == "" || os.Getenv("OOXML_FIXTURES_ROOT") == "" {
		return false
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return false
	}
	var pin referencePin
	if json.Unmarshal(data, &pin) != nil {
		return false
	}
	// loadReferencePin verifies exact HEAD, manifest, tracked bytes and clean
	// checkout before inventory or any selected case runs.
	return batch2PinMatches(pin) && batch2RootAllowed(os.Getenv("OOXML_FIXTURES_ROOT")) == nil
}

type batch2Ledger struct {
	SelectedIDs   []string `json:"selectedIds"`
	ExpandedCases int      `json:"expandedCases"`
	CompiledSteps int      `json:"compiledSteps"`
}
type batch2Workflows struct {
	Workflows []struct {
		ID            string `json:"id"`
		Feature       string `json:"feature"`
		ExpandedCases int    `json:"expandedCases"`
	} `json:"workflows"`
}

// Compile every selected case against its immutable feature and ledger owner.
// Pickle identities include the Examples-row line, not merely the scenario ID.
func batch2Inventory() (map[caseID]expectedCase, []map[string]any, error) {
	blob, err := os.ReadFile(testutil.ReferencePath("ledgers", "package-alignment.json"))
	if err != nil {
		return nil, nil, err
	}
	var ledger batch2Ledger
	if err = json.Unmarshal(blob, &ledger); err != nil {
		return nil, nil, err
	}
	if len(ledger.SelectedIDs) != 20 || ledger.ExpandedCases != 27 || ledger.CompiledSteps != 127 {
		return nil, nil, fmt.Errorf("batch-2 ledger counts changed")
	}
	registryData, err := os.ReadFile(testutil.ReferencePath("ledgers", "workflows.json"))
	if err != nil {
		return nil, nil, err
	}
	var registry batch2Workflows
	if err = json.Unmarshal(registryData, &registry); err != nil {
		return nil, nil, err
	}
	registryOwners := map[string]string{}
	registryCounts := map[string]int{}
	for _, workflow := range registry.Workflows {
		if _, exists := registryOwners[workflow.ID]; exists {
			return nil, nil, fmt.Errorf("duplicate workflow ledger ID %s", workflow.ID)
		}
		registryOwners[workflow.ID] = workflow.Feature
		registryCounts[workflow.ID] = workflow.ExpandedCases
	}
	expectedIDs := map[string]string{}
	for path, ids := range batch2Paths {
		for _, id := range ids {
			if expectedIDs[id] != "" {
				return nil, nil, fmt.Errorf("duplicate batch-2 owner %s", id)
			}
			expectedIDs[id] = path
		}
	}
	if len(expectedIDs) != 20 {
		return nil, nil, fmt.Errorf("batch-2 source map must own twenty IDs")
	}
	present := map[string]bool{}
	for _, id := range ledger.SelectedIDs {
		if present[id] || expectedIDs[id] == "" || registryOwners[id] != expectedIDs[id] || registryCounts[id] <= 0 {
			return nil, nil, fmt.Errorf("unexpected batch-2 ledger ID/owner %s", id)
		}
		present[id] = true
	}
	nextIndex := 0
	next := func() string { nextIndex++; return fmt.Sprint(nextIndex) }
	expected := map[caseID]expectedCase{}
	inventory := []map[string]any{}
	idsSeen := map[string]int{}
	steps := 0
	for relative, selected := range batch2Paths {
		path := testutil.ReferencePath(filepath.FromSlash(relative))
		f, err := os.Open(path)
		if err != nil {
			return nil, nil, err
		}
		doc, parseErr := gherkin.ParseGherkinDocument(f, next)
		closeErr := f.Close()
		if parseErr != nil {
			return nil, nil, parseErr
		}
		if closeErr != nil {
			return nil, nil, closeErr
		}
		if doc.Feature == nil {
			return nil, nil, fmt.Errorf("empty batch-2 feature %s", relative)
		}
		lines := map[string]int{}
		ownership := map[string]string{}
		var collect func([]*messages.FeatureChild) error
		collect = func(children []*messages.FeatureChild) error {
			for _, child := range children {
				if child.Rule != nil {
					for _, nested := range child.Rule.Children {
						if nested.Scenario != nil {
							if err := collect([]*messages.FeatureChild{{Scenario: nested.Scenario}}); err != nil {
								return err
							}
						}
					}
					continue
				}
				scenario := child.Scenario
				if scenario == nil {
					continue
				}
				var id string
				for _, tag := range scenario.Tags {
					for _, want := range selected {
						if tag.Name == want {
							if id != "" {
								return fmt.Errorf("multiple selected batch-2 IDs")
							}
							id = want
						}
					}
				}
				if id == "" {
					continue
				}
				if ownership[scenario.Id] != "" || idsSeen[id] != 0 {
					return fmt.Errorf("duplicate scenario identity %s", id)
				}
				ownership[scenario.Id] = id
				lines[scenario.Id] = int(scenario.Location.Line)
				for _, examples := range scenario.Examples {
					for _, row := range examples.TableBody {
						lines[row.Id] = int(row.Location.Line)
					}
				}
			}
			return nil
		}
		if err = collect(doc.Feature.Children); err != nil {
			return nil, nil, err
		}
		for _, p := range gherkin.Pickles(*doc, path, next) {
			if len(p.AstNodeIds) == 0 {
				return nil, nil, fmt.Errorf("pickle without source identity in %s", relative)
			}
			id := ownership[p.AstNodeIds[0]]
			if id == "" {
				continue
			}
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			if line < 1 || len(p.Steps) == 0 {
				return nil, nil, fmt.Errorf("invalid batch-2 row %s", id)
			}
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			if _, ok := expected[key]; ok {
				return nil, nil, fmt.Errorf("duplicate batch-2 case %+v", key)
			}
			expected[key] = expectedCase{caseID: key, Name: p.Name, Steps: len(p.Steps)}
			inventory = append(inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "sealed batch-2 candidate"})
			steps += len(p.Steps)
			idsSeen[id]++
		}
	}
	for id := range expectedIDs {
		if idsSeen[id] == 0 || idsSeen[id] != registryCounts[id] {
			return nil, nil, fmt.Errorf("batch-2 ID/count mismatch %s: got %d want %d", id, idsSeen[id], registryCounts[id])
		}
	}
	if len(idsSeen) != 20 || len(expected) != 27 || steps != 127 {
		return nil, nil, fmt.Errorf("batch-2 compiled %d IDs/%d cases/%d steps (want 20/27/127)", len(idsSeen), len(expected), steps)
	}
	for _, id := range ledger.SelectedIDs {
		if !strings.HasPrefix(id, "@id-") {
			return nil, nil, fmt.Errorf("bad batch-2 identity %s", id)
		}
	}
	return expected, inventory, nil
}

// Root identity is checked independently of Git commit/seal: an arbitrary
// clean clone of those bytes must not activate the selected lane.
func TestBatch2CandidateRootIdentity(t *testing.T) {
	if err := batch2RootAllowed(batch2Root); err != nil {
		t.Fatal(err)
	}
	alias := filepath.Join(t.TempDir(), "reviewed-root")
	if err := os.Symlink(batch2Root, alias); err != nil {
		t.Fatal(err)
	}
	if err := batch2RootAllowed(alias); err != nil {
		t.Fatalf("realpath alias rejected: %v", err)
	}
	if err := batch2RootAllowed(filepath.Join("..", "references", "fixtures-ooxml")); err == nil {
		t.Fatal("default reference checkout admitted")
	}
	if err := batch2RootAllowed(t.TempDir()); err == nil {
		t.Fatal("unreviewed temporary root admitted")
	}
}

// Fault controls alter only private parser inputs or private expected results.
// They never change the sealed reference checkout, its ledger or fixture bytes.
func TestBatch2CandidateSelectionFaults(t *testing.T) {
	if !batch2Candidate() {
		t.Skip("sealed batch-2 candidate only")
	}
	loadReferencePin(t)
	expected, _, err := batch2Inventory()
	if err != nil {
		t.Fatal(err)
	}
	if len(expected) != 27 {
		t.Fatal("missing selected cases")
	}
	// Exercise every selected Examples-row identity independently. A dropped,
	// duplicated, redirected or non-passing result cannot borrow another row's
	// evidence, including the three already selected by the historical runner.
	for selected, want := range expected {
		fixture := func(key caseID, name, status string, count int) []byte {
			step := reportStep{Name: "observed"}
			step.Result.Status = status
			feature := reportFeature{URI: key.File, Elements: []reportElement{{Name: name, Line: key.Line, Type: "scenario", Steps: make([]reportStep, count)}}}
			feature.Elements[0].Tags = append(feature.Elements[0].Tags, struct {
				Name string `json:"name"`
			}{Name: key.ID})
			for i := range feature.Elements[0].Steps {
				feature.Elements[0].Steps[i] = step
			}
			data, _ := json.Marshal([]reportFeature{feature})
			return data
		}
		one := map[caseID]expectedCase{selected: want}
		good := fixture(selected, want.Name, "passed", want.Steps)
		if err := reconcile(one, good); err != nil {
			t.Fatal(err)
		}
		if err := reconcile(expected, good); err == nil {
			t.Fatal("dropped cases accepted")
		}
		for _, status := range []string{"failed", "skipped", "undefined", "pending"} {
			if err := reconcile(one, fixture(selected, want.Name, status, want.Steps)); err == nil {
				t.Fatalf("%s step accepted for %+v", status, selected)
			}
		}
		var duplicate []reportFeature
		if err := json.Unmarshal(good, &duplicate); err != nil {
			t.Fatal(err)
		}
		duplicate[0].Elements = append(duplicate[0].Elements, duplicate[0].Elements[0])
		bytes, _ := json.Marshal(duplicate)
		if err := reconcile(one, bytes); err == nil {
			t.Fatalf("duplicate result accepted for %+v", selected)
		}
		badKey := selected
		badKey.Line++
		if err := reconcile(one, fixture(badKey, want.Name, "passed", want.Steps)); err == nil {
			t.Fatalf("wrong row line accepted for %+v", selected)
		}
		if err := reconcile(one, fixture(selected, want.Name+" changed", "passed", want.Steps)); err == nil {
			t.Fatalf("wrong name accepted for %+v", selected)
		}
		if err := reconcile(one, fixture(selected, want.Name, "passed", want.Steps-1)); err == nil {
			t.Fatalf("missing step accepted for %+v", selected)
		}
	}
}
