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

const formattingCommit = "f174097e68e4133ebbfb58daf0f336480d75f210"
const formattingManifest = "3285d9184e599a16a29d459bfc4d0b7e7f5def8bdd141a37b80a0736aebf6d9e"
const formattingFixtureID = "fixture-84e7a8a3681d0e43d4c8a29c37e68386721dfb5df13a36b8bd17a3ebc1fb1fb2"
const formattingFixtureFile = "fixtures/pptx/formatting/formatting-84e7a8a3681d.pptx"

type formattingRecord struct {
	ID, Feature, FixtureID, Kind, Operation, Target string
	Expected                                        json.RawMessage
	ChangedMembers                                  []string
}

func formattingCandidate() bool {
	if batch2RootAllowed(os.Getenv("OOXML_FIXTURES_ROOT")) != nil || os.Getenv("OOXML_REFERENCE_PIN") == "" {
		return false
	}
	b, e := os.ReadFile(os.Getenv("OOXML_REFERENCE_PIN"))
	if e != nil {
		return false
	}
	var p referencePin
	if json.Unmarshal(b, &p) != nil {
		return false
	}
	return p.Schema == 2 && p.Commit == formattingCommit && p.Manifest == formattingManifest && p.Tag == "candidate-pptx-formatting" && p.TagObject == ""
}
func formattingRecords() (map[string]formattingRecord, error) {
	b, e := os.ReadFile(testutil.ReferencePath("ledgers", "pptx-formatting.json"))
	if e != nil {
		return nil, e
	}
	var ledger struct {
		Records []formattingRecord `json:"records"`
	}
	if e = json.Unmarshal(b, &ledger); e != nil {
		return nil, e
	}
	if len(ledger.Records) != 20 {
		return nil, fmt.Errorf("formatting records %d", len(ledger.Records))
	}
	out := map[string]formattingRecord{}
	for _, r := range ledger.Records {
		if !strings.HasPrefix(r.ID, "@id-pptx-formatting-") || r.FixtureID != formattingFixtureID || !strings.HasPrefix(r.Feature, "workflows/pptx/") || r.Kind == "" || r.Operation == "" || len(r.Expected) == 0 || out[r.ID].ID != "" {
			return nil, fmt.Errorf("invalid formatting record %s", r.ID)
		}
		out[r.ID] = r
	}
	return out, nil
}
func formattingInventory() (map[caseID]expectedCase, []map[string]any, error) {
	records, e := formattingRecords()
	if e != nil {
		return nil, nil, e
	}
	owners := map[string]bool{}
	for _, r := range records {
		owners[r.Feature] = true
	}
	expected := map[caseID]expectedCase{}
	inventory := []map[string]any{}
	seen := map[string]int{}
	steps := 0
	serial := 0
	next := func() string { serial++; return fmt.Sprint(serial) }
	for relative := range owners {
		path := testutil.ReferencePath(filepath.FromSlash(relative))
		f, e := os.Open(path)
		if e != nil {
			return nil, nil, e
		}
		doc, parseErr := gherkin.ParseGherkinDocument(f, next)
		_ = f.Close()
		if parseErr != nil {
			return nil, nil, parseErr
		}
		if doc.Feature == nil {
			return nil, nil, fmt.Errorf("empty formatting feature %s", relative)
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
				return nil, nil, fmt.Errorf("formatting case without source ID")
			}
			id := names[p.AstNodeIds[0]]
			if id == "" {
				continue
			}
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			if line < 1 || len(p.Steps) == 0 {
				return nil, nil, fmt.Errorf("invalid formatting case %s", id)
			}
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			if _, dupe := expected[key]; dupe {
				return nil, nil, fmt.Errorf("duplicate formatting case %s", id)
			}
			expected[key] = expectedCase{caseID: key, Name: p.Name, Steps: len(p.Steps)}
			inventory = append(inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "sealed PPTX formatting candidate"})
			seen[id]++
			steps += len(p.Steps)
		}
	}
	if len(expected) != 20 || len(seen) != 20 || steps != 182 {
		return nil, nil, fmt.Errorf("formatting compiled %d cases %d IDs %d steps", len(expected), len(seen), steps)
	}
	for id := range records {
		if seen[id] != 1 {
			return nil, nil, fmt.Errorf("formatting ID %s expanded %d", id, seen[id])
		}
	}
	return expected, inventory, nil
}
