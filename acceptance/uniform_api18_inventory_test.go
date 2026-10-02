package acceptance

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/testutil"
)

const uniformLedgerSHA256 = "0d2a46e89f92274e3f4304564b718d1517ea5709863b4319963d29cba8d8add5"

type uniformLedger struct {
	Selected []string `json:"selectedScenarioIds"`
	Counts   struct {
		IDs        int `json:"ids"`
		Cases      int `json:"cases"`
		AfterSteps int `json:"afterSteps"`
	} `json:"counts"`
	Files []struct {
		Path        string `json:"path"`
		AfterSHA256 string `json:"afterSha256"`
		Scenarios   []struct {
			ID    string `json:"id"`
			After []struct {
				Name  string `json:"name"`
				Steps []struct {
					Text     string          `json:"text"`
					Argument json.RawMessage `json:"argument"`
				} `json:"steps"`
			} `json:"after"`
		} `json:"scenarios"`
	} `json:"files"`
}

func readUniformLedger() (uniformLedger, error) {
	var l uniformLedger
	b, err := os.ReadFile(testutil.ReferencePath("ledgers", "uniform-api18.json"))
	if err != nil {
		return l, err
	}
	if digest := sha256.Sum256(b); hex.EncodeToString(digest[:]) != uniformLedgerSHA256 {
		return l, fmt.Errorf("uniform fixture ledger provenance drift")
	}
	if err := json.Unmarshal(b, &l); err != nil {
		return l, err
	}
	return l, nil
}

func uniformSelection() (map[string]bool, error) {
	l, e := readUniformLedger()
	if e != nil {
		return nil, e
	}
	if l.Counts.IDs != 18 || l.Counts.Cases != 59 || l.Counts.AfterSteps != 279 || len(l.Selected) != 18 {
		return nil, fmt.Errorf("uniform ledger counts drift")
	}
	ids := map[string]bool{}
	for _, id := range l.Selected {
		if ids[id] {
			return nil, fmt.Errorf("duplicate %s", id)
		}
		ids[id] = true
	}
	return ids, nil
}
func uniformInventory() (map[caseID]expectedCase, []map[string]any, error) {
	ids, e := uniformSelection()
	if e != nil {
		return nil, nil, e
	}
	ledger, e := readUniformLedger()
	if e != nil {
		return nil, nil, e
	}
	result := map[caseID]expectedCase{}
	rows := []map[string]any{}
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	counts := map[string]int{}
	total := 0
	steps := 0
	if len(ledger.Files) != 2 || ledger.Files[0].Path != "workflows/xlsx/formula-references.feature" || ledger.Files[1].Path != "workflows/xml/editing.feature" {
		return nil, nil, fmt.Errorf("uniform fixture feature paths drift")
	}
	for _, sealed := range ledger.Files {
		path := testutil.ReferencePath(filepath.FromSlash(sealed.Path))
		data, err := os.ReadFile(path)
		if err != nil {
			return nil, nil, err
		}
		if digest := sha256.Sum256(data); hex.EncodeToString(digest[:]) != sealed.AfterSHA256 {
			return nil, nil, fmt.Errorf("uniform feature provenance drift: %s", sealed.Path)
		}
		f, err := os.Open(path)
		if err != nil {
			return nil, nil, err
		}
		doc, err := gherkin.ParseGherkinDocument(f, next)
		f.Close()
		if err != nil {
			return nil, nil, err
		}
		owners := map[string]string{}
		lines := map[string]int{}
		var walk func([]*messages.FeatureChild)
		walk = func(children []*messages.FeatureChild) {
			for _, child := range children {
				if child.Rule != nil {
					for _, v := range child.Rule.Children {
						if v.Scenario != nil {
							walk([]*messages.FeatureChild{{Scenario: v.Scenario}})
						}
					}
					continue
				}
				s := child.Scenario
				if s == nil {
					continue
				}
				for _, tag := range s.Tags {
					if ids[tag.Name] {
						owners[s.Id] = tag.Name
						lines[s.Id] = int(s.Location.Line)
						for _, ex := range s.Examples {
							for _, row := range ex.TableBody {
								lines[row.Id] = int(row.Location.Line)
							}
						}
					}
				}
			}
		}
		walk(doc.Feature.Children)
		sealedRows := map[string][][]string{}
		for _, scenario := range sealed.Scenarios {
			if !ids[scenario.ID] {
				return nil, nil, fmt.Errorf("unexpected sealed ID %s", scenario.ID)
			}
			for _, after := range scenario.After {
				key := scenario.ID + "\x00" + after.Name
				var texts []string
				for _, step := range after.Steps {
					if len(step.Argument) != 0 && string(step.Argument) != "null" {
						return nil, nil, fmt.Errorf("unexpected sealed step argument %s", key)
					}
					texts = append(texts, step.Text)
				}
				sealedRows[key] = append(sealedRows[key], texts)
			}
		}
		for _, p := range gherkin.Pickles(*doc, path, next) {
			if len(p.AstNodeIds) == 0 {
				continue
			}
			id := owners[p.AstNodeIds[0]]
			if id == "" {
				continue
			}
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			if line <= 0 || len(p.Steps) == 0 {
				return nil, nil, fmt.Errorf("invalid uniform case %s", id)
			}
			sealedKey := id + "\x00" + p.Name
			choices := sealedRows[sealedKey]
			matched := -1
			for j, want := range choices {
				if len(want) != len(p.Steps) {
					continue
				}
				valid := true
				for i, step := range p.Steps {
					if step.Text != want[i] || step.Argument != nil {
						valid = false
						break
					}
				}
				if valid {
					matched = j
					break
				}
			}
			if matched < 0 {
				return nil, nil, fmt.Errorf("uniform exact row/step drift: %s", sealedKey)
			}
			choices = append(choices[:matched], choices[matched+1:]...)
			if len(choices) == 0 {
				delete(sealedRows, sealedKey)
			} else {
				sealedRows[sealedKey] = choices
			}
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			if _, exists := result[key]; exists {
				return nil, nil, fmt.Errorf("duplicate case %s", id)
			}
			result[key] = expectedCase{caseID: key, Name: p.Name, Steps: len(p.Steps)}
			rows = append(rows, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "sealed uniform-api18 candidate"})
			counts[id]++
			total++
			steps += len(p.Steps)
		}
		if len(sealedRows) != 0 {
			return nil, nil, fmt.Errorf("uncompiled sealed uniform rows in %s: %d", sealed.Path, len(sealedRows))
		}
	}
	if len(counts) != 18 || total != 59 || steps != 279 {
		return nil, nil, fmt.Errorf("uniform compiled %d IDs %d cases %d steps", len(counts), total, steps)
	}
	return result, rows, nil
}
