package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/testutil"
)

const pptxManipulationCommit = "5dc02ae970e2eaf5ed3e261324ae3961e08bb1c9"
const pptxManipulationManifest = "70e44ae170ee39175f943bbd1c1a07d64d726bbdcb53507f7be95eec7b21b900"
const pptxManipulationFixture = "fixture-5ad4b68acc5a926c020dbc94154c95fa6c40c7353a55b2f82504879b392e9a45"
const pptxManipulationFixtureFile = "fixtures/pptx/manipulation/manipulation-5ad4b68acc5a.pptx"

type pptxManipulationRecord struct {
	ID, Kind, Feature, FixtureID, Target, Operation, AdditionPolicy string
	Expected                                                        json.RawMessage
	ChangedMembers                                                  []string
}
type pptxManipulationLedger struct {
	Records []pptxManipulationRecord `json:"records"`
}

func pptxManipulationCandidate() bool {
	if batch2RootAllowed(os.Getenv("OOXML_FIXTURES_ROOT")) != nil || os.Getenv("OOXML_REFERENCE_PIN") == "" {
		return false
	}
	b, err := os.ReadFile(os.Getenv("OOXML_REFERENCE_PIN"))
	if err != nil {
		return false
	}
	var pin referencePin
	if json.Unmarshal(b, &pin) != nil {
		return false
	}
	return pin.Schema == 2 && pin.Commit == pptxManipulationCommit && pin.Manifest == pptxManipulationManifest && pin.Tag == "candidate-pptx-manipulation" && pin.TagObject == ""
}
func pptxManipulationRecords() (map[string]pptxManipulationRecord, error) {
	b, err := os.ReadFile(testutil.ReferencePath("ledgers", "pptx-manipulation.json"))
	if err != nil {
		return nil, err
	}
	var ledger pptxManipulationLedger
	if err := json.Unmarshal(b, &ledger); err != nil {
		return nil, err
	}
	if len(ledger.Records) != 20 {
		return nil, fmt.Errorf("PPTX candidate ledger has %d records", len(ledger.Records))
	}
	out := map[string]pptxManipulationRecord{}
	for _, r := range ledger.Records {
		if !strings.HasPrefix(r.ID, "@id-pptx-manipulation-") || r.FixtureID != pptxManipulationFixture || !strings.HasPrefix(r.Feature, "workflows/pptx/") || r.Kind == "" || r.Operation == "" || len(r.Expected) == 0 || out[r.ID].ID != "" {
			return nil, fmt.Errorf("invalid PPTX candidate record %s", r.ID)
		}
		out[r.ID] = r
	}
	return out, nil
}

// The newer sealed PPTX checkout includes the batch-2 spelling/layout of
// historical shared cases. It does not select or rerun batch-2's added lane.
func postBatch2Reference() bool {
	return batch2Candidate() || pptxManipulationCandidate() || formattingCandidate() || retainedCandidate() || tableCandidate() || uniformAPI18Candidate()
}

func pptxManipulationSelectedID(id string) bool {
	return strings.HasPrefix(id, "@id-pptx-manipulation-")
}
func pptxManipulationInventory() (map[caseID]expectedCase, []map[string]any, error) {
	records, err := pptxManipulationRecords()
	if err != nil {
		return nil, nil, err
	}
	owners := map[string][]string{}
	for id, r := range records {
		owners[r.Feature] = append(owners[r.Feature], id)
	}
	expected := map[caseID]expectedCase{}
	inventory := []map[string]any{}
	seen := map[string]int{}
	steps := 0
	serial := 0
	next := func() string { serial++; return fmt.Sprint(serial) }
	for relative := range owners {
		path := testutil.ReferencePath(filepath.FromSlash(relative))
		f, err := os.Open(path)
		if err != nil {
			return nil, nil, err
		}
		doc, parseErr := gherkin.ParseGherkinDocument(f, next)
		_ = f.Close()
		if parseErr != nil {
			return nil, nil, parseErr
		}
		if doc.Feature == nil {
			return nil, nil, fmt.Errorf("empty PPTX feature %s", relative)
		}
		names := map[string]string{}
		lines := map[string]int{}
		var walk func([]*messages.FeatureChild)
		walk = func(children []*messages.FeatureChild) {
			for _, child := range children {
				if child.Rule != nil {
					for _, nested := range child.Rule.Children {
						if nested.Scenario != nil {
							walk([]*messages.FeatureChild{{Scenario: nested.Scenario}})
						}
					}
					continue
				}
				if child.Scenario == nil {
					continue
				}
				sc := child.Scenario
				for _, tag := range sc.Tags {
					if r, ok := records[tag.Name]; ok && r.Feature == relative {
						names[sc.Id] = tag.Name
						lines[sc.Id] = int(sc.Location.Line)
						for _, examples := range sc.Examples {
							for _, row := range examples.TableBody {
								lines[row.Id] = int(row.Location.Line)
							}
						}
					}
				}
			}
		}
		walk(doc.Feature.Children)
		for _, p := range gherkin.Pickles(*doc, path, next) {
			if len(p.AstNodeIds) == 0 {
				return nil, nil, fmt.Errorf("PPTX case without source ID")
			}
			id := names[p.AstNodeIds[0]]
			if id == "" {
				continue
			}
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			if line < 1 || len(p.Steps) == 0 {
				return nil, nil, fmt.Errorf("invalid PPTX case %s", id)
			}
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			if _, duplicate := expected[key]; duplicate {
				return nil, nil, fmt.Errorf("duplicate PPTX case %s", id)
			}
			expected[key] = expectedCase{caseID: key, Name: p.Name, Steps: len(p.Steps)}
			inventory = append(inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "sealed PPTX manipulation candidate"})
			seen[id]++
			steps += len(p.Steps)
		}
	}
	if len(expected) != 20 || len(seen) != 20 || steps != 179 {
		return nil, nil, fmt.Errorf("PPTX candidate compiled %d cases/%d IDs/%d steps, want 20/20/179", len(expected), len(seen), steps)
	}
	for id := range records {
		if seen[id] != 1 {
			return nil, nil, fmt.Errorf("PPTX candidate ID %s expanded %d times", id, seen[id])
		}
	}
	return expected, inventory, nil
}
