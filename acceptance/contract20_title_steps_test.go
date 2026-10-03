package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"
	"unicode/utf16"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type contract20TitleWorld struct {
	recipe                     string
	temp                       string
	source, original, saved    []byte
	baseFixture                []byte
	before, after, postSuccess map[string][]byte
	graph                      packaging.Graph
	session                    *presentation.EditSession
	anchor                     *presentation.ContractTitleAnchor
	unaffected                 *presentation.ContractTitleAnchor
	operands                   struct {
		start, end       int
		old, replacement string
	}
	result bool
	err    error
	prior  string
}

func (w *contract20TitleWorld) input(recipeID string) error {
	sealed, err := os.ReadFile(testutil.ReferencePath("ledgers", "contract20-recipes.json"))
	if err != nil {
		return err
	}
	h := sha256.Sum256(sealed)
	if hex.EncodeToString(h[:]) != contract20Artifacts["ledgers/contract20-recipes.json"] {
		return fmt.Errorf("PPTX recipe seal differs")
	}
	var recipe struct {
		PPTX []struct {
			ID, BaseFixtureID string
			Operations        []struct{ Kind, Part, Before, After, Value string }
		}
	}
	if err := json.Unmarshal(sealed, &recipe); err != nil {
		return err
	}
	found := 0
	for _, r := range recipe.PPTX {
		if r.ID != recipeID {
			continue
		}
		found++
		fixture, err := testutil.LookupFixture(r.BaseFixtureID)
		if err != nil {
			return err
		}
		base, err := os.ReadFile(fixture)
		if err != nil {
			return err
		}
		w.baseFixture = bytes.Clone(base)
		members, err := blankMemberPayloads(base)
		if err != nil {
			return err
		}
		for _, op := range r.Operations {
			switch op.Kind {
			case "replace-literal-once":
				original, ok := members[op.Part]
				if !ok || strings.Count(string(original), op.Before) != 1 {
					return fmt.Errorf("nonunique title recipe anchor %s", op.Part)
				}
				members[op.Part] = []byte(strings.Replace(string(original), op.Before, op.After, 1))
			case "add-literal-member":
				if _, exists := members[op.Part]; exists {
					return fmt.Errorf("existing added title member")
				}
				members[op.Part] = []byte(op.Value)
			default:
				return fmt.Errorf("unsupported title recipe operation %q", op.Kind)
			}
		}
		w.source, err = contract20PackMembers(members)
		if err != nil {
			return err
		}
		w.original = bytes.Clone(w.source)
		w.before, err = blankMemberPayloads(w.source)
		if err != nil {
			return err
		}
		w.graph, err = contract20TableGraph(w.source)
		if err != nil {
			return err
		}
		w.session, err = presentation.OpenEditing(w.source, packaging.Limits{})
		if err != nil {
			return err
		}
		w.prior = filepath.Join(w.temp, "prior.pptx")
		if err = os.WriteFile(w.prior, []byte("prior-destination"), 0600); err != nil {
			return err
		}
		w.recipe = recipeID
	}
	if found != 1 {
		return fmt.Errorf("title recipe %s count %d", recipeID, found)
	}
	return nil
}
func (w *contract20TitleWorld) find() error {
	var err error
	w.anchor, err = w.session.FindContractTitle("ppt/slides/slide1.xml")
	return err
}
func (w *contract20TitleWorld) plain() error {
	if w.anchor == nil || w.anchor.Text() != "Frankenstein" {
		return fmt.Errorf("plain title %q", w.anchor.Text())
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) fragmented() error {
	if w.anchor == nil || w.anchor.Text() != "Frankenstein" {
		return fmt.Errorf("fragmented title %q", w.anchor.Text())
	}
	want := []presentation.ContractTitleFragment{{Text: "Fran", Attributes: map[string]string{"b": "1"}}, {Text: "ken", Attributes: map[string]string{"i": "1"}}, {Text: "stein", Attributes: map[string]string{"u": "sng"}}}
	if !reflect.DeepEqual(w.anchor.Fragments(), want) {
		return fmt.Errorf("original fragments %+v", w.anchor.Fragments())
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) breakField() error {
	if w.anchor == nil || w.anchor.Text() != "Chapter\n7 Notes" {
		return fmt.Errorf("break/field title %q", w.anchor.Text())
	}
	want := []presentation.ContractTitleFragment{{Text: "Chapter", Attributes: map[string]string{"b": "1"}}, {Text: "\n", Attributes: map[string]string{"lang": "en-US"}}, {Text: "7", Attributes: map[string]string{"i": "1"}}, {Text: " Notes", Attributes: map[string]string{"u": "sng"}}}
	if !reflect.DeepEqual(w.anchor.Fragments(), want) {
		return fmt.Errorf("break/field fragments %+v", w.anchor.Fragments())
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) record() error {
	if len(w.before) == 0 || len(w.graph.Edges) == 0 || len(w.source) == 0 {
		return fmt.Errorf("source not recorded")
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) unchanged() error {
	if !bytes.Equal(w.source, w.original) {
		return fmt.Errorf("caller archive changed")
	}
	members, err := blankMemberPayloads(w.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(members, w.before) {
		return fmt.Errorf("source members changed")
	}
	graph, err := contract20TableGraph(w.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(graph, w.graph) {
		return fmt.Errorf("source graph changed")
	}
	b, err := os.ReadFile(w.prior)
	if err != nil || string(b) != "prior-destination" {
		return fmt.Errorf("prior destination changed: %v", err)
	}
	return nil
}
func (w *contract20TitleWorld) replace(start, end int, old, value string) error {
	if w.anchor == nil {
		return fmt.Errorf("missing title anchor")
	}
	units := utf16.Encode([]rune(w.anchor.Text()))
	if start < 0 || end > len(units) || start > end || (old != "" && string(utf16.Decode(units[start:end])) != old) {
		return fmt.Errorf("unexpected UTF-16 operand")
	}
	w.operands = struct {
		start, end       int
		old, replacement string
	}{start, end, old, value}
	w.result, w.err = w.session.ReplaceContractTitleSpan(w.anchor, start, end, value)
	if w.err != nil {
		return w.err
	}
	return nil
}
func (w *contract20TitleWorld) attempt(start, end int, value string) error {
	w.result, w.err = w.session.ReplaceContractTitleSpan(w.anchor, start, end, value)
	return nil
}
func (w *contract20TitleWorld) refused(kind string) error {
	var refusal *packaging.Refusal
	if w.result || !errors.As(w.err, &refusal) || refusal.Kind != kind {
		return fmt.Errorf("expected %s/no change, got %v/%v", kind, w.result, w.err)
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) observeAfterRefusal() error {
	t, err := w.session.FindContractTitle("ppt/slides/slide1.xml")
	if err != nil {
		return err
	}
	if t.Text() != "Chapter\n7 Notes" {
		return fmt.Errorf("refused title no longer readable %q", t.Text())
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) save() error {
	path := filepath.Join(w.temp, "edited.pptx")
	if _, err := w.session.SaveAs(path); err != nil {
		return err
	}
	var err error
	w.saved, err = os.ReadFile(path)
	if err != nil {
		return err
	}
	w.after, err = blankMemberPayloads(w.saved)
	if err != nil {
		return err
	}
	if _, err = presentation.OpenEditing(w.saved, packaging.Limits{}); err != nil {
		return err
	}
	return nil
}
func (w *contract20TitleWorld) graphPreserved() error {
	graph, err := contract20TableGraph(w.saved)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(graph, w.graph) {
		return fmt.Errorf("saved OPC graph changed")
	}
	return nil
}
func (w *contract20TitleWorld) savedTitle(text string) error {
	if len(w.after) == 0 {
		return fmt.Errorf("no saved member state")
	}
	if got, err := oracleContract20Title(w.after["ppt/slides/slide1.xml"]); err != nil || got != text {
		return fmt.Errorf("saved title %q expected %q: %v", got, text, err)
	}
	s, err := presentation.OpenEditing(w.saved, packaging.Limits{})
	if err != nil {
		return err
	}
	t, err := s.FindContractTitle("ppt/slides/slide1.xml")
	if err != nil {
		return err
	}
	if t.Text() != text {
		return fmt.Errorf("reopened title %q", t.Text())
	}
	return nil
}
func (w *contract20TitleWorld) custody() error {
	if err := w.unchanged(); err != nil {
		return err
	}
	if w.recipe == "fragmented-title" && w.operands.old != "" {
		units := utf16.Encode([]rune(w.anchor.Text()))
		if w.operands.start < 0 || w.operands.end > len(units) || string(utf16.Decode(units[w.operands.start:w.operands.end])) != w.operands.old || w.operands.replacement != "iend" {
			return fmt.Errorf("cross-run operands changed")
		}
	}
	if len(w.after) != len(w.before) {
		return fmt.Errorf("member inventory changed")
	}
	for part, original := range w.before {
		saved, ok := w.after[part]
		if !ok {
			return fmt.Errorf("member missing: %s", part)
		}
		if part != "ppt/slides/slide1.xml" && !bytes.Equal(saved, original) {
			return fmt.Errorf("unrelated payload changed: %s", part)
		}
	}
	return w.graphPreserved()
}
func (w *contract20TitleWorld) changed() error {
	var changes []string
	for part, before := range w.before {
		if !bytes.Equal(before, w.after[part]) {
			changes = append(changes, part)
		}
	}
	if !reflect.DeepEqual(changes, []string{"ppt/slides/slide1.xml"}) {
		return fmt.Errorf("changed members %q", changes)
	}
	return w.custody()
}
func (w *contract20TitleWorld) literalSpans() error {
	before, after := w.before["ppt/slides/slide1.xml"], w.after["ppt/slides/slide1.xml"]
	const old = `<a:p><a:r><a:rPr b="1"/><a:t>Fran</a:t></a:r><a:r><a:rPr i="1"/><a:t>ken</a:t></a:r><a:r><a:rPr u="sng"/><a:t>stein</a:t></a:r></a:p>`
	const expected = `<a:p><a:r><a:rPr b="1"/><a:t>Fr</a:t></a:r><a:r><a:rPr b="1"/><a:t>iend</a:t></a:r><a:r><a:rPr u="sng"/><a:t>stein</a:t></a:r></a:p>`
	if strings.Count(string(before), old) != 1 || strings.Count(string(after), expected) != 1 || !bytes.Equal(after, []byte(strings.Replace(string(before), old, expected, 1))) {
		return fmt.Errorf("changed title lexical mask mismatch")
	}
	return nil
}
func (w *contract20TitleWorld) runProperties() error {
	got, err := oracleContract20TitleRuns(w.after["ppt/slides/slide1.xml"])
	if err != nil {
		return err
	}
	want := []presentation.ContractTitleFragment{{Text: "Fr", Attributes: map[string]string{"b": "1"}}, {Text: "iend", Attributes: map[string]string{"b": "1"}}, {Text: "stein", Attributes: map[string]string{"u": "sng"}}}
	if !reflect.DeepEqual(got, want) {
		return fmt.Errorf("independent saved runs %+v", got)
	}
	return w.literalSpans()
}
func (w *contract20TitleWorld) saveFaults() error {
	var err error
	w.unaffected, err = w.session.FindContractTitle("ppt/slides/slide1.xml")
	if err != nil || w.unaffected.Text() != "Friendstein" {
		return fmt.Errorf("fresh held edited title %q: %v", w.unaffected.Text(), err)
	}
	// Both snapshots are delivered through the actual current session. The
	// earlier successful archive is not a proxy for state after failed saves.
	capture := func(filename string) (map[string][]byte, packaging.Graph, error) {
		path := filepath.Join(w.temp, filename)
		if _, err := w.session.SaveAs(path); err != nil {
			return nil, packaging.Graph{}, err
		}
		archive, err := os.ReadFile(path)
		if err != nil {
			return nil, packaging.Graph{}, err
		}
		parts, err := blankMemberPayloads(archive)
		if err != nil {
			return nil, packaging.Graph{}, err
		}
		graph, err := contract20TableGraph(archive)
		return parts, graph, err
	}
	priorMembers, priorGraph, err := capture("pre-fault-current.pptx")
	if err != nil {
		return err
	}
	if err := contract20SaveFaults(filepath.Join(w.temp, "fault-paths"), func(path string) error { _, err := w.session.SaveAs(path); return err }); err != nil {
		return err
	}
	afterMembers, afterGraph, err := capture("post-fault-current.pptx")
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(afterMembers, priorMembers) || !reflect.DeepEqual(afterGraph, priorGraph) || !reflect.DeepEqual(afterGraph, w.graph) {
		return fmt.Errorf("failed SaveAs changed live session members or graph")
	}
	for part, original := range w.before {
		if part != "ppt/slides/slide1.xml" && !bytes.Equal(afterMembers[part], original) {
			return fmt.Errorf("failed SaveAs changed unrelated part %s", part)
		}
	}
	t, err := w.session.FindContractTitle("ppt/slides/slide1.xml")
	if err != nil {
		return err
	}
	if t.Text() != "Friendstein" {
		return fmt.Errorf("held edited title lost after failed saves")
	}
	// Failed saves must not consume an issued fresh anchor. An exact
	// replacement no-op exercises its issuance and unchanged part fingerprint.
	changed, err := w.session.ReplaceContractTitleSpan(w.unaffected, 0, 1, "F")
	if err != nil || changed {
		return fmt.Errorf("held title unusable after save fault: %v/%v", changed, err)
	}
	return w.custody()
}
func (w *contract20TitleWorld) recordPostSuccess() error {
	path := filepath.Join(w.temp, "post-success.pptx")
	if _, err := w.session.SaveAs(path); err != nil {
		return err
	}
	archive, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	w.postSuccess, err = blankMemberPayloads(archive)
	if err != nil {
		return err
	}
	return w.unchanged()
}
func (w *contract20TitleWorld) stale() error {
	if w.result || w.err == nil {
		return fmt.Errorf("stale edit accepted")
	}
	var refusal *packaging.Refusal
	if !errors.As(w.err, &refusal) || refusal.Kind != "PPTX_STALE_ANCHOR" {
		return fmt.Errorf("wrong stale refusal %v", w.err)
	}
	return w.postSuccessCustody()
}
func (w *contract20TitleWorld) postSuccessCustody() error {
	if err := w.unchanged(); err != nil {
		return err
	}
	path := filepath.Join(w.temp, "post-success-verify.pptx")
	if _, err := w.session.SaveAs(path); err != nil {
		return err
	}
	archive, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	members, err := blankMemberPayloads(archive)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(members, w.postSuccess) {
		return fmt.Errorf("stale refusal changed post-success members")
	}
	return nil
}
func (w *contract20TitleWorld) newlyIssued() error {
	fresh, err := w.session.FindContractTitle("ppt/slides/slide1.xml")
	if err != nil {
		return err
	}
	if fresh.Text() != "Creature" {
		return fmt.Errorf("fresh title %q", fresh.Text())
	}
	return w.custody()
}

// Independent decoder: no production lossless XML parser or title API.
func oracleContract20Title(data []byte) (string, error) {
	runs, err := oracleContract20TitleRuns(data)
	if err != nil {
		return "", err
	}
	var out strings.Builder
	for _, r := range runs {
		out.WriteString(r.Text)
	}
	return out.String(), nil
}
func oracleContract20TitleRuns(data []byte) ([]presentation.ContractTitleFragment, error) {
	d := xml.NewDecoder(bytes.NewReader(data))
	var stack []xml.Name
	var shapeDepth, bodyDepth, paraDepth, runDepth, textDepth = -1, -1, -1, -1, -1
	selected := false
	count := 0
	var fragments, observed []presentation.ContractTitleFragment
	var current presentation.ContractTitleFragment
	p := func(s string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: s} }
	a := func(s string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: s} }
	for {
		token, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch v := token.(type) {
		case xml.StartElement:
			parent := xml.Name{}
			if len(stack) > 0 {
				parent = stack[len(stack)-1]
			}
			stack = append(stack, v.Name)
			depth := len(stack) - 1
			switch {
			case v.Name == p("sp") && shapeDepth < 0:
				shapeDepth = depth
				selected = false
				bodyDepth = -1
				paraDepth = -1
				fragments = nil
			case shapeDepth >= 0 && v.Name == p("ph"):
				for _, attr := range v.Attr {
					if attr.Name.Space == "" && attr.Name.Local == "type" && (attr.Value == "title" || attr.Value == "ctrTitle") {
						selected = true
					}
				}
			case shapeDepth >= 0 && v.Name == p("txBody") && parent == p("sp") && depth == shapeDepth+1:
				bodyDepth = depth
			case bodyDepth >= 0 && v.Name == a("p") && parent == p("txBody") && depth == bodyDepth+1 && paraDepth < 0:
				paraDepth = depth
			case paraDepth >= 0 && (v.Name == a("r") || v.Name == a("fld") || v.Name == a("br")) && parent == a("p") && depth == paraDepth+1:
				runDepth = depth
				current = presentation.ContractTitleFragment{Attributes: map[string]string{}}
				if v.Name == a("br") {
					current.Text = "\n"
				}
			case runDepth >= 0 && v.Name == a("rPr") && depth == runDepth+1:
				for _, attr := range v.Attr {
					if attr.Name.Space == "" {
						current.Attributes[attr.Name.Local] = attr.Value
					}
				}
			case runDepth >= 0 && v.Name == a("t") && depth == runDepth+1:
				textDepth = depth
			}
		case xml.CharData:
			if textDepth == len(stack)-1 && textDepth >= 0 {
				current.Text += string(v)
			}
		case xml.EndElement:
			depth := len(stack) - 1
			if depth < 0 || stack[depth] != v.Name {
				return nil, fmt.Errorf("oracle XML stack mismatch")
			}
			if depth == textDepth {
				textDepth = -1
			}
			if depth == runDepth {
				fragments = append(fragments, current)
				runDepth = -1
			}
			if depth == paraDepth {
				paraDepth = -1
			}
			if depth == bodyDepth {
				bodyDepth = -1
			}
			if depth == shapeDepth {
				if selected {
					count++
					if count > 1 {
						return nil, fmt.Errorf("oracle duplicate title")
					}
					observed = append([]presentation.ContractTitleFragment(nil), fragments...)
				}
				shapeDepth = -1
			}
			stack = stack[:depth]
		}
	}
	if count != 1 {
		return nil, fmt.Errorf("oracle title count %d", count)
	}
	return observed, nil
}

func TestContract20TitleSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct{ id, recipe string }{
		{"@id-pptx-readable-unsupported-topology", "break-field-title"},
		{"@id-pptx-cross-run-replace", "fragmented-title"},
		{"@id-pptx-stale-anchor-refusal", "plain-title"},
	} {
		var key caseID
		var want expectedCase
		count := 0
		for k, v := range inventory {
			if k.ID == tc.id {
				key, want = k, v
				count++
			}
		}
		if count != 1 {
			t.Fatalf("title case %s has %d rows", tc.id, count)
		}
		t.Run(strings.TrimPrefix(tc.id, "@id-"), func(t *testing.T) {
			state := &contract20TitleWorld{temp: t.TempDir()}
			exact := func(sc *godog.ScenarioContext, text string, action func() error) {
				sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
			}
			init := func(sc *godog.ScenarioContext) {
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
					*state = contract20TitleWorld{temp: t.TempDir()}
					return ctx, nil
				})
				switch tc.recipe {
				case "break-field-title":
					exact(sc, `the contract20 derived PPTX recipe "break-field-title" is created in memory from its sealed fixture`, func() error { return state.input(tc.recipe) })
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, state.record)
					exact(sc, `the production text inspector reads the uniquely selected first title paragraph and issues its owned anchor`, state.find)
					exact(sc, `the exact paragraph text equals JSON "Chapter\n7 Notes"`, state.breakField)
					exact(sc, `the exact visible fragments and properties equal JSON [{"text":"Chapter","attributes":{"b":"1"}},{"text":"\n","attributes":{"lang":"en-US"}},{"text":"7","attributes":{"i":"1"}},{"text":" Notes","attributes":{"u":"sng"}}]`, state.breakField)
					exact(sc, `the production anchored editor attempts JSON "Chapter" to JSON "Section" through that anchor`, func() error { return state.attempt(0, 7, "Section") })
					exact(sc, `the operation returns exactly typed refusal category "PPTX_UNSUPPORTED_TEXT_TOPOLOGY" with no changed result`, func() error { return state.refused("PPTX_UNSUPPORTED_TEXT_TOPOLOGY") })
					exact(sc, `every original member, relationship, caller archive and prior destination remains unchanged`, state.unchanged)
					exact(sc, `the held paragraph remains readable and a subsequent no-op observation leaves the source unchanged`, state.observeAfterRefusal)
				case "fragmented-title":
					exact(sc, `the contract20 derived PPTX recipe "fragmented-title" is created in memory from its sealed fixture`, func() error { return state.input(tc.recipe) })
					exact(sc, `the unique first title paragraph reads JSON "Frankenstein" with source runs JSON [{"text":"Fran","attributes":{"b":"1"}},{"text":"ken","attributes":{"i":"1"}},{"text":"stein","attributes":{"u":"sng"}}]`, func() error {
						if err := state.find(); err != nil {
							return err
						}
						return state.fragmented()
					})
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, state.record)
					exact(sc, `the production anchored editor replaces UTF-16 substring [2,7) JSON "anken" with JSON "iend"`, func() error { return state.replace(2, 7, "anken", "iend") })
					exact(sc, `the result is saved to a distinct new path and independently parsed and reopened`, state.save)
					exact(sc, `the source fixture, caller archive and operation operands remain unchanged`, state.unchanged)
					exact(sc, `only the sealed original member and lexical span allowances differ; all unrelated member payloads remain literal`, state.literalSpans)
					exact(sc, `all saved OPC relationships and content types resolve with original identities preserved`, state.graphPreserved)
					exact(sc, `refusals and save faults publish no partial destination and leave the session and held unaffected targets usable`, state.saveFaults)
					exact(sc, `the reopened title text equals JSON "Friendstein"`, func() error { return state.savedTitle("Friendstein") })
					exact(sc, `the exact saved runs equal JSON [{"text":"Fr","attributes":{"b":"1"}},{"text":"iend","attributes":{"b":"1"}},{"text":"stein","attributes":{"u":"sng"}}]`, state.runProperties)
					exact(sc, `the exact changed member set is JSON ["ppt/slides/slide1.xml"] without additions or removals`, state.changed)
					exact(sc, `only consumed run text spans and the inherited replacement run may differ; every unrelated selected paragraph property and sibling span remains literal`, state.literalSpans)
					exact(sc, `custom/data.bin equals JSON "keep-me-safe" and every other member payload remains unchanged`, func() error {
						if !bytes.Equal(state.after["custom/data.bin"], []byte("keep-me-safe")) {
							return fmt.Errorf("custom binary changed")
						}
						return state.custody()
					})
				case "plain-title":
					exact(sc, `the contract20 derived PPTX recipe "plain-title" is created in memory from its sealed fixture`, func() error { return state.input(tc.recipe) })
					exact(sc, `the production inspector holds the unique first title paragraph anchor reading JSON "Frankenstein"`, func() error {
						if err := state.find(); err != nil {
							return err
						}
						return state.plain()
					})
					exact(sc, `the production anchored editor replaces JSON "Frankenstein" with JSON "Creature" through that anchor`, func() error { return state.replace(0, 12, "Frankenstein", "Creature") })
					exact(sc, `the complete post-success package member state and prior destination are recorded`, state.recordPostSuccess)
					exact(sc, `the same old anchor attempts replacement JSON "Frankenstein" with JSON "Monster"`, func() error { return state.attempt(0, 12, "Monster") })
					exact(sc, `the operation returns exactly typed refusal category "PPTX_STALE_ANCHOR" and no changed result`, state.stale)
					exact(sc, `the complete post-success member state, original caller input and prior destination remain unchanged`, state.postSuccessCustody)
					exact(sc, `saving and independently reopening the retained session reads JSON "Creature"`, func() error {
						if err := state.save(); err != nil {
							return err
						}
						return state.savedTitle("Creature")
					})
					exact(sc, `a newly issued target for the first title remains usable and original relationships and unrelated members retain custody`, state.newlyIssued)
				}
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-title", ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: tc.id, Strict: true, Concurrency: 1}}
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
				if err := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(tc.id, "@id-")+".cucumber.json"), output.Bytes(), 0600); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
}
