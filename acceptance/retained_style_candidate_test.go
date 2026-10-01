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

const retainedCommit = "7f9a3469d335b01d30b6bbd8acdecc490bc9f0d0"
const retainedManifest = "1522f8f9c77e36038ed4510bdb99578059dc893301ff53db214fbcc56ddf03f2"

type retainedRecord struct {
	ID, Feature, FixtureID, Kind, Format, Target string
	Patch, Expected                              json.RawMessage
	Mask                                         json.RawMessage
	ChangedMembers                               []string
}

func retainedCandidate() bool {
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
	return p.Schema == 2 && p.Commit == retainedCommit && p.Manifest == retainedManifest && p.Tag == "candidate-retained-style-word" && p.TagObject == ""
}
func retainedRecords() (map[string]retainedRecord, error) {
	b, e := os.ReadFile(testutil.ReferencePath("ledgers", "retained-style-word.json"))
	if e != nil {
		return nil, e
	}
	var l struct {
		Records []retainedRecord `json:"records"`
	}
	if e = json.Unmarshal(b, &l); e != nil {
		return nil, e
	}
	if len(l.Records) != 40 {
		return nil, fmt.Errorf("retained records %d", len(l.Records))
	}
	out := map[string]retainedRecord{}
	formats := map[string]int{}
	for _, r := range l.Records {
		if !(strings.HasPrefix(r.ID, "@id-pptx-retained-") && r.Format == "pptx" || strings.HasPrefix(r.ID, "@id-docx-retained-") && r.Format == "docx") || !strings.HasPrefix(r.Feature, "workflows/"+r.Format+"/") || !strings.HasPrefix(r.FixtureID, "fixture-") || r.Kind == "" || len(r.Patch) == 0 || len(r.Expected) == 0 || out[r.ID].ID != "" {
			return nil, fmt.Errorf("invalid retained record %s", r.ID)
		}
		out[r.ID] = r
		formats[r.Format]++
	}
	if formats["pptx"] != 20 || formats["docx"] != 20 {
		return nil, fmt.Errorf("retained formats %v", formats)
	}
	return out, nil
}
func retainedInventory() (map[caseID]expectedCase, []map[string]any, error) {
	records, e := retainedRecords()
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
	stepsByFormat := map[string]int{}
	serial := 0
	next := func() string { serial++; return fmt.Sprint(serial) }
	for relative := range owners {
		path := testutil.ReferencePath(filepath.FromSlash(relative))
		f, err := os.Open(path)
		if err != nil {
			return nil, nil, err
		}
		doc, pe := gherkin.ParseGherkinDocument(f, next)
		_ = f.Close()
		if pe != nil {
			return nil, nil, pe
		}
		if doc.Feature == nil {
			return nil, nil, fmt.Errorf("empty retained feature %s", relative)
		}
		names := map[string]string{}
		lines := map[string]int{}
		var walk func([]*messages.FeatureChild)
		walk = func(children []*messages.FeatureChild) {
			for _, child := range children {
				if child.Rule != nil {
					for _, n := range child.Rule.Children {
						if n.Scenario != nil {
							walk([]*messages.FeatureChild{{Scenario: n.Scenario}})
						}
					}
					continue
				}
				if child.Scenario == nil {
					continue
				}
				s := child.Scenario
				for _, tag := range s.Tags {
					if r, ok := records[tag.Name]; ok && r.Feature == relative {
						names[s.Id] = tag.Name
						lines[s.Id] = int(s.Location.Line)
						for _, examples := range s.Examples {
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
				return nil, nil, fmt.Errorf("retained case without source ID")
			}
			id := names[p.AstNodeIds[0]]
			if id == "" {
				continue
			}
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			if line < 1 || len(p.Steps) == 0 {
				return nil, nil, fmt.Errorf("invalid retained pickle %s", id)
			}
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			if _, ok := expected[key]; ok {
				return nil, nil, fmt.Errorf("duplicate retained case %s", id)
			}
			expected[key] = expectedCase{caseID: key, Name: p.Name, Steps: len(p.Steps)}
			inventory = append(inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "sealed retained style/Word candidate"})
			seen[id]++
			stepsByFormat[records[id].Format] += len(p.Steps)
		}
	}
	if len(expected) != 40 || len(seen) != 40 || stepsByFormat["pptx"] != 201 || stepsByFormat["docx"] != 200 {
		return nil, nil, fmt.Errorf("retained compiled %d cases %d IDs %v steps", len(expected), len(seen), stepsByFormat)
	}
	for id := range records {
		if seen[id] != 1 {
			return nil, nil, fmt.Errorf("retained ID %s expanded %d", id, seen[id])
		}
	}
	return expected, inventory, nil
}
