package acceptance

import (
	"crypto/sha256"
	"encoding/hex"
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

// Contract20 is an explicit, untagged candidate. This inventory grants no
// execution credit: each selected case still needs independent Go bindings.
const contract20Commit = "f4b5be7998ae6d009429f2080b40e3d4530c3452"
const contract20Manifest = "e8a6fc090257f64d2a1487edbfea0b151ff0aff66de35d4d0d098d3b35284aa7"
const contract20LedgerHash = "1814d45d98ff0c596da99324fcaa2ec47982c89535120ba8b1919c47006a1b43"

var contract20Artifacts = map[string]string{
	"contracts/contract20.md":                 "5d8f5d19d70c2716a5d3688ff9d625ae5683175d5c88468bb3740a0e95c6edd0",
	"ledgers/contract20-recipes.json":         "9587da306418a4e67e7f23304beacdd5adec49719a7319342879bc315396c6fa",
	"ledgers/contract20-creation-policy.json": "34334c3967af72a3e335ed9ec6aa728ba822b14eada26f48b16b845c61844853",
	"ledgers/contract20.json":                 contract20LedgerHash,
}

var contract20IDs = []string{
	"@id-pptx-order-notes-read", "@id-pptx-readable-unsupported-topology", "@id-pptx-cross-run-replace", "@id-pptx-stale-anchor-refusal",
	"@id-xlsx-read-rel-linked-shared-strings", "@id-xlsx-clear-cross-sheet-caches", "@id-xlsx-prefixed-namespace-safe-edits", "@id-xlsx-phonetic-guides-excluded",
	"@id-xlsx-styled-blank-cell-editable", "@id-xlsx-refuse-shared-formula-overwrite", "@id-xlsx-refuse-array-formula-overwrite", "@id-xlsx-array-input-refusal",
	"@id-xlsx-cache-scope-opaque-parts", "@id-pptx-create-new-minimal", "@id-pptx-create-anchor-preservation", "@id-pptx-create-refusals",
	"@id-pptx-table-roundtrip-geometry", "@id-pptx-table-formatting", "@id-pptx-table-stale-handle", "@id-pptx-table-atomic-refusals",
}

type contract20Case struct {
	Name  string `json:"name"`
	Steps []struct {
		Text     string          `json:"text"`
		Argument json.RawMessage `json:"argument"`
	} `json:"steps"`
}
type contract20Ledger struct {
	SourceRevision  string   `json:"sourceRevision"`
	ExecutionCredit bool     `json:"executionCredit"`
	Selected        []string `json:"selectedScenarioIds"`
	Counts          struct {
		IDs         int `json:"ids"`
		Cases       int `json:"cases"`
		BeforeSteps int `json:"beforeSteps"`
		AfterSteps  int `json:"afterSteps"`
	} `json:"counts"`
	Files []struct {
		Path      string `json:"path"`
		AfterSHA  string `json:"afterSha256"`
		Scenarios []struct {
			ID    string           `json:"id"`
			After []contract20Case `json:"after"`
		} `json:"scenarios"`
	} `json:"files"`
}

func contract20SelectedID(id string) bool {
	for _, selected := range contract20IDs {
		if id == selected {
			return true
		}
	}
	return false
}

func contract20Candidate() bool {
	if batch2RootAllowed(os.Getenv("OOXML_FIXTURES_ROOT")) != nil || os.Getenv("OOXML_REFERENCE_PIN") == "" {
		return false
	}
	data, err := os.ReadFile(os.Getenv("OOXML_REFERENCE_PIN"))
	if err != nil {
		return false
	}
	var pin referencePin
	if json.Unmarshal(data, &pin) != nil {
		return false
	}
	return pin.Schema == 2 && pin.Commit == contract20Commit && pin.Manifest == contract20Manifest && pin.Tag == "candidate-contract20" && pin.TagObject == "" && pin.Pack == "" && pin.Assets == 373 && pin.Facts == 149 && pin.Workflows == 424 && pin.Cases == 911
}

func readContract20Ledger() (contract20Ledger, error) {
	var l contract20Ledger
	for path, want := range contract20Artifacts {
		b, err := os.ReadFile(testutil.ReferencePath(filepath.FromSlash(path)))
		if err != nil {
			return l, err
		}
		h := sha256.Sum256(b)
		if hex.EncodeToString(h[:]) != want {
			return l, fmt.Errorf("Contract20 artifact drift: %s", path)
		}
		if path == "ledgers/contract20.json" && json.Unmarshal(b, &l) != nil {
			return l, fmt.Errorf("invalid Contract20 ledger")
		}
	}
	if l.SourceRevision != "0531afe2f0bddb50879cf7f11d5f7035e3fe374f" || l.ExecutionCredit || l.Counts.IDs != 20 || l.Counts.Cases != 22 || l.Counts.BeforeSteps != 74 || l.Counts.AfterSteps != 223 || len(l.Files) != 6 || len(l.Selected) != 20 {
		return l, fmt.Errorf("Contract20 ledger scope drift")
	}
	for i, id := range contract20IDs {
		if l.Selected[i] != id {
			return l, fmt.Errorf("Contract20 selection drift: %d", i)
		}
	}
	return l, nil
}

// Compile only selected IDs, compare the actual expanded step text (not merely
// counts), and refuse any execution-credit assertion in the shared ledger.
func contract20Inventory() (map[caseID]expectedCase, error) {
	l, err := readContract20Ledger()
	if err != nil {
		return nil, err
	}
	var workflows struct {
		Workflows []struct {
			ID        string `json:"id"`
			Feature   string `json:"feature"`
			Consumers map[string]struct {
				Status   string          `json:"status"`
				Evidence json.RawMessage `json:"evidence"`
			} `json:"consumers"`
		} `json:"workflows"`
	}
	b, err := os.ReadFile(testutil.ReferencePath("ledgers", "workflows.json"))
	if err != nil {
		return nil, err
	}
	if err = json.Unmarshal(b, &workflows); err != nil {
		return nil, err
	}
	owners := map[string]string{}
	for _, row := range workflows.Workflows {
		owners[row.ID] = row.Feature
	}
	selected := map[string]bool{}
	for _, id := range l.Selected {
		selected[id] = true
	}
	for _, row := range workflows.Workflows {
		if selected[row.ID] {
			for _, language := range []string{"bun", "go", "python"} {
				c, ok := row.Consumers[language]
				if !ok || c.Status != "planned" || string(c.Evidence) != "[]" {
					return nil, fmt.Errorf("Contract20 forged consumer credit: %s/%s", row.ID, language)
				}
			}
		}
	}
	result := map[caseID]expectedCase{}
	seen := map[string]int{}
	totalSteps := 0
	serial := 0
	next := func() string { serial++; return fmt.Sprint(serial) }
	for _, f := range l.Files {
		if !filepath.IsLocal(f.Path) || !strings.HasPrefix(f.Path, "workflows/") {
			return nil, fmt.Errorf("unsafe Contract20 feature path")
		}
		path := testutil.ReferencePath(filepath.FromSlash(f.Path))
		content, err := os.ReadFile(path)
		if err != nil {
			return nil, err
		}
		h := sha256.Sum256(content)
		if hex.EncodeToString(h[:]) != f.AfterSHA {
			return nil, fmt.Errorf("Contract20 feature drift %s", f.Path)
		}
		file, err := os.Open(path)
		if err != nil {
			return nil, err
		}
		doc, parseErr := gherkin.ParseGherkinDocument(file, next)
		_ = file.Close()
		if parseErr != nil {
			return nil, parseErr
		}
		if doc.Feature == nil {
			return nil, fmt.Errorf("empty Contract20 feature %s", f.Path)
		}
		ids := map[string]string{}
		lines := map[string]int{}
		sealed := map[string][]contract20Case{}
		for _, s := range f.Scenarios {
			if !selected[s.ID] || owners[s.ID] != f.Path || len(s.After) == 0 {
				return nil, fmt.Errorf("unexpected Contract20 ID %s", s.ID)
			}
			sealed[s.ID] = append([]contract20Case(nil), s.After...)
		}
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
				s := child.Scenario
				for _, tag := range s.Tags {
					if _, ok := sealed[tag.Name]; ok {
						ids[s.Id] = tag.Name
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
		for _, p := range gherkin.Pickles(*doc, path, next) {
			if len(p.AstNodeIds) == 0 {
				return nil, fmt.Errorf("case missing source ID")
			}
			id := ids[p.AstNodeIds[0]]
			if id == "" {
				continue
			}
			line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
			if line <= 0 {
				return nil, fmt.Errorf("invalid Contract20 case row %s", id)
			}
			matches := sealed[id]
			found := -1
			for i, row := range matches {
				if row.Name != p.Name || len(row.Steps) != len(p.Steps) {
					continue
				}
				same := true
				for j, step := range p.Steps {
					if step.Text != row.Steps[j].Text || step.Argument != nil && (len(row.Steps[j].Argument) == 0 || string(row.Steps[j].Argument) == "null") || step.Argument == nil && len(row.Steps[j].Argument) != 0 && string(row.Steps[j].Argument) != "null" {
						same = false
						break
					}
				}
				if same {
					found = i
					break
				}
			}
			if found == -1 {
				return nil, fmt.Errorf("Contract20 case/step drift: %s %s", id, p.Name)
			}
			sealed[id] = append(matches[:found], matches[found+1:]...)
			key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
			if _, exists := result[key]; exists {
				return nil, fmt.Errorf("duplicate Contract20 case %s", id)
			}
			result[key] = expectedCase{caseID: key, Name: p.Name, Steps: len(p.Steps)}
			seen[id]++
			totalSteps += len(p.Steps)
		}
		for id, left := range sealed {
			if len(left) != 0 {
				return nil, fmt.Errorf("uncompiled Contract20 case %s", id)
			}
		}
	}
	if len(seen) != 20 || len(result) != 22 || totalSteps != 223 {
		return nil, fmt.Errorf("Contract20 compiled %d IDs/%d cases/%d steps", len(seen), len(result), totalSteps)
	}
	return result, nil
}

func TestContract20CandidateInventory(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate pin/root not selected")
	}
	loadReferencePin(t) // verifies exact clean HEAD and all tracked bytes before parsing
	if _, err := contract20Inventory(); err != nil {
		t.Fatal(err)
	}
	for _, id := range contract20IDs {
		if id == "" {
			t.Fatal("empty selected ID")
		}
	}
	// Physical manifest custody is a separate admission check, not a runtime receipt.
	TestPinnedReferenceDistribution(t)
}
