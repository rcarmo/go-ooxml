package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type contract20NotesReadWorld struct {
	sources          map[string][]byte
	originals        map[string][]byte
	members          map[string]map[string][]byte
	graphs           map[string]packaging.Graph
	ordered, missing []presentation.ContractSlideNotes
}

func (w *contract20NotesReadWorld) input() error {
	b, err := os.ReadFile(testutil.ReferencePath("ledgers", "contract20-recipes.json"))
	if err != nil {
		return err
	}
	h := sha256.Sum256(b)
	if hex.EncodeToString(h[:]) != contract20Artifacts["ledgers/contract20-recipes.json"] {
		return fmt.Errorf("notes recipe seal differs")
	}
	var recipes struct {
		PPTX []struct {
			ID, BaseFixtureID string
			Operations        []struct {
				Kind, Part, Before, After string
				OriginalCount             int
				OneBasedOrder             []int
			}
		}
	}
	if err = json.Unmarshal(b, &recipes); err != nil {
		return err
	}
	w.sources = map[string][]byte{}
	w.originals = map[string][]byte{}
	w.members = map[string]map[string][]byte{}
	w.graphs = map[string]packaging.Graph{}
	for _, id := range []string{"ordered-notes", "missing-notes"} {
		var found bool
		for _, recipe := range recipes.PPTX {
			if recipe.ID != id {
				continue
			}
			if found {
				return fmt.Errorf("duplicate notes recipe %s", id)
			}
			found = true
			path, e := testutil.LookupFixture(recipe.BaseFixtureID)
			if e != nil {
				return e
			}
			archive, e := os.ReadFile(path)
			if e != nil {
				return e
			}
			parts, e := blankMemberPayloads(archive)
			if e != nil {
				return e
			}
			for _, op := range recipe.Operations {
				original, ok := parts[op.Part]
				if !ok {
					return fmt.Errorf("recipe part missing %s", op.Part)
				}
				switch op.Kind {
				case "replace-literal-once":
					if strings.Count(string(original), op.Before) != 1 {
						return fmt.Errorf("recipe anchor ambiguous %s", op.Part)
					}
					parts[op.Part] = []byte(strings.Replace(string(original), op.Before, op.After, 1))
				case "reorder-sldId-elements":
					if op.OriginalCount != 5 || !reflect.DeepEqual(op.OneBasedOrder, []int{5, 1, 4, 2, 3}) {
						return fmt.Errorf("unsealed slide order")
					}
					doc, e := losslessxml.Parse(original)
					if e != nil {
						return e
					}
					var ranges [][2]int
					for _, n := range doc.Elements() {
						if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldId"}) {
							continue
						}
						owner, ok := n.Parent()
						if !ok || owner.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldIdLst"}) {
							return fmt.Errorf("slide id outside list")
						}
						a, z := n.SourceRange()
						ranges = append(ranges, [2]int{a, z})
					}
					if len(ranges) != op.OriginalCount {
						return fmt.Errorf("slide list count %d", len(ranges))
					}
					for j := 1; j < len(ranges); j++ {
						if len(bytes.TrimSpace(original[ranges[j-1][1]:ranges[j][0]])) != 0 {
							return fmt.Errorf("slide ID separators not whitespace")
						}
					}
					var replacement bytes.Buffer
					used := map[int]bool{}
					for _, position := range op.OneBasedOrder {
						if position < 1 || position > len(ranges) || used[position] {
							return fmt.Errorf("slide reorder is not a permutation")
						}
						used[position] = true
						r := ranges[position-1]
						replacement.Write(original[r[0]:r[1]])
					}
					parts[op.Part] = append(append(bytes.Clone(original[:ranges[0][0]]), replacement.Bytes()...), original[ranges[len(ranges)-1][1]:]...)
				default:
					return fmt.Errorf("unsupported notes recipe operation %s", op.Kind)
				}
			}
			w.sources[id], e = contract20PackMembers(parts)
			if e != nil {
				return e
			}
			w.originals[id] = bytes.Clone(w.sources[id])
			w.members[id], e = blankMemberPayloads(w.sources[id])
			if e != nil {
				return e
			}
			w.graphs[id], e = contract20TableGraph(w.sources[id])
			if e != nil {
				return e
			}
		}
		if !found {
			return fmt.Errorf("notes recipe missing %s", id)
		}
	}
	return nil
}
func (w *contract20NotesReadWorld) record() error {
	for _, id := range []string{"ordered-notes", "missing-notes"} {
		if len(w.members[id]) == 0 || len(w.sources[id]) == 0 || len(w.graphs[id].Edges) == 0 {
			return fmt.Errorf("notes source not recorded %s", id)
		}
	}
	return nil
}
func (w *contract20NotesReadWorld) read() error {
	var err error
	w.ordered, err = presentation.ReadContractSlideNotes(w.sources["ordered-notes"])
	if err != nil {
		return err
	}
	w.missing, err = presentation.ReadContractSlideNotes(w.sources["missing-notes"])
	return err
}
func (w *contract20NotesReadWorld) titles() error {
	want := []string{"Chapter 5", "Frankenstein Lecture Series", "Chapter 4", "Chapter 2", "Chapter 3"}
	got := []string{}
	for _, slide := range w.ordered {
		got = append(got, slide.Title)
	}
	if !reflect.DeepEqual(got, want) {
		return fmt.Errorf("relationship-ordered titles %q", got)
	}
	// Independently prove the ordered list matches literal slide ID and OPC edge
	// targets, not merely the production reader's reported part identities.
	parts := w.members["ordered-notes"]
	var registry struct {
		IDs []struct {
			ID  string `xml:"id,attr"`
			RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
		} `xml:"sldIdLst>sldId"`
	}
	if err := xml.Unmarshal(parts["ppt/presentation.xml"], &registry); err != nil {
		return err
	}
	if len(registry.IDs) != 5 {
		return fmt.Errorf("slide IDs %d", len(registry.IDs))
	}
	var relationships struct {
		Edges []struct {
			ID     string `xml:"Id,attr"`
			Type   string `xml:"Type,attr"`
			Target string `xml:"Target,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal(parts["ppt/_rels/presentation.xml.rels"], &relationships); err != nil {
		return err
	}
	edges := map[string]string{}
	for _, edge := range relationships.Edges {
		if edge.Type != packaging.RelTypeSlide {
			continue
		}
		if edge.ID == "" || edge.Target == "" || edges[edge.ID] != "" || strings.HasPrefix(edge.Target, "/") || strings.Contains(edge.Target, "..") {
			return fmt.Errorf("invalid independent slide edge %q/%q", edge.ID, edge.Target)
		}
		edges[edge.ID] = path.Join("ppt", edge.Target)
	}
	if len(edges) != 5 {
		return fmt.Errorf("independent slide edge count %d", len(edges))
	}
	if len(w.ordered) != len(registry.IDs) {
		return fmt.Errorf("observed slide count %d", len(w.ordered))
	}
	seen := map[string]bool{}
	for i, id := range registry.IDs {
		rid := id.RID
		if rid == "" {
			return fmt.Errorf("slide RID missing")
		}
		part := edges[rid]
		if part == "" || part != w.ordered[i].Part || seen[part] {
			return fmt.Errorf("slide %d relationship drift: %s/%s", i, part, w.ordered[i].Part)
		}
		seen[part] = true
		oracle, err := contract20OraclePlaceholderText(parts[part], "", "title", "ctrTitle")
		if err != nil || oracle != w.ordered[i].Title {
			return fmt.Errorf("independent title %s: observed %q, oracle %q: %v", part, w.ordered[i].Title, oracle, err)
		}
	}
	return nil
}
func (w *contract20NotesReadWorld) texts() error {
	if len(w.ordered) != 5 {
		return fmt.Errorf("notes count %d", len(w.ordered))
	}
	wantFirst := "Key themes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12"
	wantLast := "Compare to Prometheus myth"
	if !w.ordered[0].HasNotes || w.ordered[0].Text != wantFirst || !w.ordered[4].HasNotes || w.ordered[4].Text != wantLast {
		return fmt.Errorf("first/last notes %q/%q", w.ordered[0].Text, w.ordered[4].Text)
	}
	for _, i := range []int{0, 4} {
		part := ""
		for _, edge := range w.graphs["ordered-notes"].Edges {
			if edge.Source == w.ordered[i].Part && edge.Type == packaging.RelTypeNotesSlide && !edge.External {
				if part != "" {
					return fmt.Errorf("ambiguous independent notes edges for slide %d", i)
				}
				part = edge.ResolvedPart
			}
		}
		if part == "" {
			return fmt.Errorf("no notes edge for slide %d", i)
		}
		oracle, err := contract20OraclePlaceholderText(w.members["ordered-notes"][part], "\n", "body")
		if err != nil || oracle != w.ordered[i].Text {
			return fmt.Errorf("independent notes %s: observed %q, oracle %q: %v", part, w.ordered[i].Text, oracle, err)
		}
	}
	return nil
}
func (w *contract20NotesReadWorld) absent() error {
	if len(w.missing) == 0 || w.missing[0].HasNotes || w.missing[0].Text != "" {
		return fmt.Errorf("missing notes reported %+v", w.missing)
	}
	graph := w.graphs["missing-notes"]
	for _, e := range graph.Edges {
		if e.Source == w.missing[0].Part && e.Type == packaging.RelTypeNotesSlide {
			return fmt.Errorf("missing slide owns notes edge")
		}
	}
	return nil
}
func (w *contract20NotesReadWorld) custody() error {
	for _, id := range []string{"ordered-notes", "missing-notes"} {
		if !bytes.Equal(w.originals[id], w.sources[id]) {
			return fmt.Errorf("caller archive changed %s", id)
		}
		parts, err := blankMemberPayloads(w.sources[id])
		if err != nil {
			return err
		}
		if !reflect.DeepEqual(parts, w.members[id]) {
			return fmt.Errorf("read changed member bytes %s", id)
		}
		graph, err := contract20TableGraph(w.sources[id])
		if err != nil {
			return err
		}
		if !reflect.DeepEqual(graph, w.graphs[id]) {
			return fmt.Errorf("read changed graph %s", id)
		}
	}
	return w.absent()
}
func (w *contract20NotesReadWorld) lineCustody() error {
	expected := "Key themes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12"
	if len(w.ordered) < 1 || w.ordered[0].Text != expected {
		return fmt.Errorf("read line breaks were trimmed")
	}
	part := ""
	for _, e := range w.graphs["ordered-notes"].Edges {
		if e.Source == w.ordered[0].Part && e.Type == packaging.RelTypeNotesSlide {
			part = e.ResolvedPart
		}
	}
	data := w.members["ordered-notes"][part]
	if part == "" || !bytes.Contains(data, []byte(`<a:br/>`)) || !bytes.Contains(data, []byte(`<a:p></a:p>`)) || !bytes.Contains(data, []byte(`type="datetimeFigureOut"`)) {
		return fmt.Errorf("sealed break/blank/field input absent")
	}
	oracle, err := contract20OraclePlaceholderText(data, "\n", "body")
	if err != nil || oracle != expected || oracle != w.ordered[0].Text {
		return fmt.Errorf("independent break/blank/field text %q, observed %q: %v", oracle, w.ordered[0].Text, err)
	}
	return w.custody()
}

func TestContract20NotesReadSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	const id = "@id-pptx-order-notes-read"
	var key caseID
	var want expectedCase
	count := 0
	for k, v := range inventory {
		if k.ID == id {
			key, want = k, v
			count++
		}
	}
	if count != 1 {
		t.Fatalf("selected notes case rows %d", count)
	}
	state := &contract20NotesReadWorld{}
	exact := func(sc *godog.ScenarioContext, text string, action func() error) {
		sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
	}
	init := func(sc *godog.ScenarioContext) {
		sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
			*state = contract20NotesReadWorld{}
			return ctx, nil
		})
		exact(sc, `the contract20 derived PPTX recipes "ordered-notes" and "missing-notes" are created in memory from their sealed fixtures`, state.input)
		exact(sc, `the source member payloads, relationships and caller archives are recorded`, state.record)
		exact(sc, `the production reader inspects presentation-relationship order and notes without saving`, state.read)
		exact(sc, `the exact ordered slide titles equal JSON ["Chapter 5","Frankenstein Lecture Series","Chapter 4","Chapter 2","Chapter 3"]`, state.titles)
		exact(sc, `the first and last ordered notes texts equal JSON ["Key themes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12","Compare to Prometheus myth"]`, state.texts)
		exact(sc, `notes absence on the missing-notes first slide returns normalized JSON {"hasNotes":false,"text":""}`, state.absent)
		exact(sc, `no read or absence observation adds notes parts or relationships or changes any member or caller archive`, state.custody)
		exact(sc, `all notes line breaks are normalized to newline without evaluating fields or trimming intentional interior blank paragraphs`, state.lineCustody)
	}
	var output bytes.Buffer
	suite := godog.TestSuite{Name: "go-contract20-notes-read", ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: id, Strict: true, Concurrency: 1}}
	if code := suite.Run(); code != 0 {
		t.Errorf("Godog %d: %s", code, output.String())
	}
	var report []reportFeature
	if err := json.Unmarshal(output.Bytes(), &report); err != nil {
		t.Fatalf("report: %v: %s", err, output.String())
	}
	data, err := json.Marshal(report)
	if err != nil {
		t.Fatal(err)
	}
	if err := reconcile(map[caseID]expectedCase{key: want}, data); err != nil {
		t.Error(err)
	}
	if dir := os.Getenv("GO_CONTRACT20_PPTX_REMAINING_RECEIPTS_DIR"); dir != "" && !t.Failed() {
		if e := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(id, "@id-")+".cucumber.json"), output.Bytes(), 0600); e != nil {
			t.Fatal(e)
		}
	}
}
