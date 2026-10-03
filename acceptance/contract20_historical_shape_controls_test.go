package acceptance

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
)

func historicalControlPickle(t *testing.T, path, id, name string) *messages.Pickle {
	t.Helper()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil {
		t.Fatal(err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name != id || (name != "" && p.Name != name) {
				continue
			}
			if selected != nil {
				t.Fatalf("duplicate %s", id)
			}
			selected = p
		}
	}
	if selected == nil {
		t.Fatalf("missing %s", id)
	}
	return selected
}

func cloneHistoricalPickle(p *messages.Pickle) *messages.Pickle {
	copy := *p
	copy.Steps = make([]*messages.PickleStep, len(p.Steps))
	for i, step := range p.Steps {
		s := *step
		copy.Steps[i] = &s
	}
	return &copy
}

// Report historical inventory separately from the Contract20 candidate's
// 20/22/223 ledger: these are existing Go-owned selections, not new credit.
func TestContract20HistoricalInventoryCensus(t *testing.T) {
	loadReferencePin(t)
	expected, _, err := inventoryCases()
	if err != nil {
		t.Fatal(err)
	}
	paths := []struct {
		name, path                   string
		defaultCount, candidateCount int
	}{
		{"xml parsing", xmlParsingFeaturePath(), 5, 8},
		{"xml names", xmlNamesFeaturePath(), 1, 1},
		{"xml editing", xmlEditingFeaturePath(), 15, 16},
		{"formula references", formulaReferenceFeaturePath(), 45, 45},
	}
	for _, item := range paths {
		var ids []string
		for key := range expected {
			if key.File == filepath.ToSlash(filepath.Clean(item.path)) {
				ids = append(ids, key.ID)
			}
		}
		sort.Strings(ids)
		want := item.defaultCount
		if contract20Candidate() {
			want = item.candidateCount
		}
		if len(ids) != want {
			t.Errorf("%s historical selected cases=%d, want %d", item.name, len(ids), want)
		}
		t.Logf("historical %s selected cases=%d IDs=%q", item.name, len(ids), ids)
	}
	if contract20Candidate() {
		selected, err := contract20Inventory()
		if err != nil {
			t.Fatal(err)
		}
		if len(selected) != 22 {
			t.Fatalf("Contract20 selected cases=%d, want 22", len(selected))
		}
		t.Logf("separate Contract20 selected IDs=%d cases=%d predicates=%d", len(contract20IDs), len(selected), 223)
	}
}

func TestContract20HistoricalShapeControls(t *testing.T) {
	loadReferencePin(t)
	if !contract20Candidate() {
		if historicalAPI18Shape() != postBatch2Reference() || !historicalStepCount(&messages.Pickle{Steps: []*messages.PickleStep{{Text: "old"}}}, 1) {
			t.Fatal("default or earlier candidate shape/steps changed")
		}
		if _, err := stableID([]string{"@planned", xmlApplyEditsID}); err == nil {
			t.Fatal("default pin admitted Contract20-only historical apply-edits ID")
		}
		return
	}
	if got, err := stableID([]string{"@planned", xmlApplyEditsID}); err != nil || got != xmlApplyEditsID {
		t.Fatalf("Contract20 historical apply-edits ID=%q err=%v", got, err)
	}
	if _, err := stableID([]string{"@planned", "@id-xml-unknown-historical"}); err == nil {
		t.Fatal("admitted unknown historical ID")
	}
	if _, err := stableID([]string{"@planned", xmlApplyEditsID, xmlEntityValuesCaseID}); err == nil {
		t.Fatal("admitted multiple historical IDs")
	}
	if !historicalAPI18Shape() || postBatch2Reference() || uniformAPI18Candidate() || batch2Candidate() || xmlLexicalCandidate() {
		t.Fatal("Contract20 shape enabled unrelated execution selection")
	}
	pin, err := os.ReadFile(os.Getenv("OOXML_REFERENCE_PIN"))
	if err != nil {
		t.Fatal(err)
	}
	var original referencePin
	if err := json.Unmarshal(pin, &original); err != nil {
		t.Fatal(err)
	}
	for _, change := range []struct {
		name string
		edit func(*referencePin)
	}{
		{"wrong tag", func(p *referencePin) { p.Tag = "candidate-uniform-api18" }},
		{"wrong commit", func(p *referencePin) { p.Commit = uniformAPI18Commit }},
		{"wrong manifest", func(p *referencePin) { p.Manifest = uniformAPI18Manifest }},
		{"wrong case count", func(p *referencePin) { p.Cases++ }},
	} {
		t.Run(change.name, func(t *testing.T) {
			mutant := original
			change.edit(&mutant)
			b, err := json.Marshal(mutant)
			if err != nil {
				t.Fatal(err)
			}
			path := filepath.Join(t.TempDir(), "pin.json")
			if err := os.WriteFile(path, b, 0600); err != nil {
				t.Fatal(err)
			}
			t.Setenv("OOXML_REFERENCE_PIN", path)
			if contract20Candidate() || historicalAPI18Shape() {
				t.Fatal("wrong pin enabled shape")
			}
		})
	}
	t.Run("wrong root", func(t *testing.T) {
		t.Setenv("OOXML_FIXTURES_ROOT", t.TempDir())
		if contract20Candidate() || historicalAPI18Shape() {
			t.Fatal("wrong root enabled shape")
		}
	})

	editing := xmlEditingFeaturePath()
	p := historicalControlPickle(t, editing, childInsertionCustodyCaseID, "Insert a child with independently scoped element and attribute names")
	f, err := os.Open(editing)
	if err != nil {
		t.Fatal(err)
	}
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	_ = f.Close()
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, child := range doc.Feature.Children {
		if child.Rule == nil {
			continue
		}
		for _, member := range child.Rule.Children {
			if s := member.Scenario; s != nil && s.Name == "Insert a child with independently scoped element and attribute names" {
				found = true
				if err := guardContract20ChildInsertionTags(s); err != nil {
					t.Fatal(err)
				}
			}
		}
	}
	if !found {
		t.Fatal("canonical child insertion scenario missing")
	}
	if err := guardChildInsertionCustodyCase(childInsertionCustodyCaseID, p, 49); err != nil {
		t.Fatal(err)
	}
	for _, change := range []struct {
		name string
		edit func(*messages.Pickle)
	}{
		{"authored value", func(p *messages.Pickle) {
			p.Steps[3].Text = "the authored attribute expanded name is other/a with JSON value \"v\" and the plain grandchild text is JSON \"text\""
		}},
		{"missing authored", func(p *messages.Pickle) { p.Steps = append(p.Steps[:3], p.Steps[4:]...) }},
		{"missing uniform", func(p *messages.Pickle) { p.Steps = p.Steps[:len(p.Steps)-1] }},
		{"duplicate uniform", func(p *messages.Pickle) { p.Steps = append(p.Steps, p.Steps[len(p.Steps)-1]) }},
		{"altered uniform", func(p *messages.Pickle) { p.Steps[len(p.Steps)-1].Text += " changed" }},
		{"argument", func(p *messages.Pickle) { p.Steps[3].Argument = &messages.PickleStepArgument{} }},
		{"old prefix", func(p *messages.Pickle) { p.Steps[2].Text += " changed" }},
	} {
		t.Run(change.name, func(t *testing.T) {
			mutant := cloneHistoricalPickle(p)
			change.edit(mutant)
			if guardChildInsertionCustodyCase(childInsertionCustodyCaseID, mutant, 49) == nil {
				t.Fatal("accepted child insertion drift")
			}
		})
	}
	if guardChildInsertionCustodyCase(childInsertionCustodyCaseID, p, 50) == nil || guardChildInsertionCustodyCase("@wrong-id", p, 49) == nil {
		t.Fatal("accepted child insertion line or ID drift")
	}
	replacement := historicalControlPickle(t, editing, elementReplacementCustodyCaseID, "Replace one subtree using its surviving parent namespace scope")
	if err := guardElementReplacementCustodyCase(elementReplacementCustodyCaseID, replacement, 105); err != nil {
		t.Fatal(err)
	}
	for _, change := range []struct {
		name string
		edit func(*messages.Pickle)
	}{
		{"missing replacement uniform", func(p *messages.Pickle) { p.Steps = p.Steps[:3] }},
		{"duplicate replacement uniform", func(p *messages.Pickle) { p.Steps = append(p.Steps, p.Steps[3]) }},
		{"altered replacement uniform", func(p *messages.Pickle) { p.Steps[3].Text += " changed" }},
		{"replacement argument", func(p *messages.Pickle) { p.Steps[3].Argument = &messages.PickleStepArgument{} }},
	} {
		t.Run(change.name, func(t *testing.T) {
			mutant := cloneHistoricalPickle(replacement)
			change.edit(mutant)
			if guardElementReplacementCustodyCase(elementReplacementCustodyCaseID, mutant, 105) == nil {
				t.Fatal("accepted replacement drift")
			}
		})
	}
	formula := historicalControlPickle(t, formulaReferenceFeaturePath(), directRangeParsingCaseID, "A absolute column direct range yields bounded axis and first coordinate")
	if err := guardDirectRangeCase(directRangeParsingCaseID, formula, formulaReferenceFeaturePath()); err != nil {
		t.Fatal(err)
	}
	for _, change := range []struct {
		name string
		edit func(*messages.Pickle)
	}{
		{"missing formula tail", func(p *messages.Pickle) { p.Steps = p.Steps[:len(p.Steps)-1] }},
		{"duplicate formula tail", func(p *messages.Pickle) { p.Steps = append(p.Steps, p.Steps[len(p.Steps)-1]) }},
		{"altered formula tail", func(p *messages.Pickle) { p.Steps[len(p.Steps)-1].Text += " changed" }},
		{"formula argument", func(p *messages.Pickle) { p.Steps[len(p.Steps)-1].Argument = &messages.PickleStepArgument{} }},
		{"formula prefix argument", func(p *messages.Pickle) { p.Steps[0].Argument = &messages.PickleStepArgument{} }},
		{"formula prefix", func(p *messages.Pickle) { p.Steps[0].Text += " changed" }},
	} {
		t.Run(change.name, func(t *testing.T) {
			mutant := cloneHistoricalPickle(formula)
			change.edit(mutant)
			if guardDirectRangeCase(directRangeParsingCaseID, mutant, formulaReferenceFeaturePath()) == nil {
				t.Fatal("accepted formula drift")
			}
		})
	}
	for _, change := range []struct {
		name, before, after string
	}{
		{"profile", "@profile-lexical-snapshot-api @id-xml-go-child-insertion-custody", "@profile-changed @id-xml-go-child-insertion-custody"},
		{"ID", "@profile-lexical-snapshot-api @id-xml-go-child-insertion-custody", "@profile-lexical-snapshot-api @id-xml-changed"},
	} {
		t.Run(change.name, func(t *testing.T) {
			// The same exact-tag guard is called by canonical inventoryFeature.
			data, err := os.ReadFile(editing)
			if err != nil {
				t.Fatal(err)
			}
			old := []byte(change.before)
			if bytes.Count(data, old) != 1 {
				t.Fatal("expected one tag pair")
			}
			mutant := bytes.Replace(data, old, []byte(change.after), 1)
			path := filepath.Join(t.TempDir(), "editing.feature")
			if err := os.WriteFile(path, mutant, 0600); err != nil {
				t.Fatal(err)
			}
			f, err := os.Open(path)
			if err != nil {
				t.Fatal(err)
			}
			defer f.Close()
			n := 0
			next := func() string { n++; return fmt.Sprint(n) }
			doc, err := gherkin.ParseGherkinDocument(f, next)
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, child := range doc.Feature.Children {
				if child.Rule == nil {
					continue
				}
				for _, member := range child.Rule.Children {
					if s := member.Scenario; s != nil && s.Name == "Insert a child with independently scoped element and attribute names" {
						found = true
						if guardContract20ChildInsertionTags(s) == nil {
							t.Fatal("canonical inventory accepted profile/ID drift")
						}
					}
				}
			}
			if !found {
				t.Fatal("mutated scenario missing")
			}
		})
	}
}
