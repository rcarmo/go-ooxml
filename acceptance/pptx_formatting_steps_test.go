package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type formattingState struct {
	record               formattingRecord
	source, output       []byte
	caller, callerBefore []byte
	before, after        map[string][]byte
	session              *presentation.EditSession
	target               *presentation.FormatTarget
	refusal              error
	held                 []byte
}

func formatScalar(v any) string {
	if n, ok := v.(float64); ok {
		return strconv.FormatInt(int64(n), 10)
	}
	return fmt.Sprint(v)
}
func formatJSON(raw json.RawMessage) (map[string]any, error) {
	out := map[string]any{}
	e := json.Unmarshal(raw, &out)
	return out, e
}
func formattingSteps(sc *godog.ScenarioContext) {
	state := &formattingState{}
	sc.Before(func(ctx context.Context, scenario *godog.Scenario) (context.Context, error) {
		*state = formattingState{}
		for _, tag := range scenario.Tags {
			if strings.HasPrefix(tag.Name, "@id-pptx-formatting-") {
				records, e := formattingRecords()
				if e != nil {
					return ctx, e
				}
				state.record = records[tag.Name]
			}
		}
		return ctx, nil
	})
	sc.Step(`^retained formatting fixture fixture-([a-f0-9]{64}) has ordered slides \["Alpha","Beta","Gamma"\]$`, func(id string) error {
		if state.record.ID == "" || id != strings.TrimPrefix(formattingFixtureID, "fixture-") {
			return fmt.Errorf("formatting case or fixture ID mismatch")
		}
		b, e := os.ReadFile(testutil.ReferencePath(filepath.FromSlash(formattingFixtureFile)))
		if e != nil {
			return e
		}
		h := sha256.Sum256(b)
		if hex.EncodeToString(h[:]) != id {
			return fmt.Errorf("formatting source provenance")
		}
		state.caller = b
		state.callerBefore = append([]byte{}, b...)
		state.source = append([]byte{}, b...)
		state.before, e = pptxMembers(b)
		if e != nil {
			return e
		}
		_, _, _, titles, e := pptxSlideTitles(b)
		if e != nil {
			return e
		}
		if !reflect.DeepEqual(titles, []string{"Alpha", "Beta", "Gamma"}) {
			return fmt.Errorf("source slide titles %v", titles)
		}
		state.session, e = presentation.OpenEditing(b, packaging.Limits{})
		return e
	})
	sc.Step(`^formatting target (.+) has the exact original direct properties and lexical snapshot$`, func(target string) error {
		if state.session == nil {
			return fmt.Errorf("formatting session absent")
		}
		if target != state.record.Target {
			return fmt.Errorf("target %q want %q", target, state.record.Target)
		}
		var e error
		state.target, e = state.session.FindFormatShape("ppt/slides/slide2.xml", 4)
		return e
	})
	sc.Step(`^production PPTX formatting APIs (.+)$`, func(operation string) error {
		if operation != state.record.Operation {
			return fmt.Errorf("operation %q want %q", operation, state.record.Operation)
		}
		return state.perform()
	})
	sc.Step(`^the result or unchanged refusal session is saved to a new path and independently parsed and reopened$`, func() error { return state.save() })
	sc.Step(`^formatting target (.+) has exact saved properties JSON (.+)$`, func(target, raw string) error {
		want, e := formatJSON(state.record.Expected)
		if e != nil {
			return e
		}
		var supplied map[string]any
		if e = json.Unmarshal([]byte(raw), &supplied); e != nil {
			return e
		}
		if !reflect.DeepEqual(want, supplied) {
			return fmt.Errorf("feature expectation differs from sealed ledger")
		}
		return state.expected(want)
	})
	sc.Step(`^the exact changed original member set is (.+) with no additions or removals$`, func(raw string) error {
		var names []string
		if e := json.Unmarshal([]byte(raw), &names); e != nil {
			return e
		}
		if !reflect.DeepEqual(names, state.record.ChangedMembers) {
			return fmt.Errorf("feature changed members %v want %v", names, state.record.ChangedMembers)
		}
		return state.custody()
	})
	sc.Step(`^every original text leaf, unselected run, paragraph, shape and unrelated member retains its original bytes$`, func() error { return state.custody() })
	sc.Step(`^caller archive, original slide IDs, relationship IDs, targets, layout and notes links retain custody$`, func() error { return state.custody() })
	sc.Step(`^only the selected property span may change and all unpatched attributes and child fragments retain their exact bytes$`, func() error { return state.custody() })
	sc.Step(`^refusal reason (.+) leaves session bytes and held identities unchanged before save$`, func(reason string) error {
		if reason != "invalid-visibility" && reason != "invalid-geometry" {
			return fmt.Errorf("unknown refusal %s", reason)
		}
		return state.checkRefusal()
	})

}
func (s *formattingState) perform() error {
	part := "ppt/slides/slide2.xml"
	switch s.record.Kind {
	case "hide":
		return s.session.SetSlideVisibility(part, false)
	case "unhide":
		if e := s.session.SetSlideVisibility(part, false); e != nil {
			return e
		}
		return s.session.SetSlideVisibility(part, true)
	case "visibility-order":
		if e := s.session.SetSlideVisibility(part, false); e != nil {
			return e
		}
		return s.session.ReorderSlides([]int{2, 0, 1})
	case "visibility-refusal":
		s.held = append([]byte{}, s.source...)
		s.refusal = s.session.SetSlideVisibilityValue(part, "hidden")
		return nil
	case "move":
		x, y := int64(914400), int64(1828800)
		return s.session.SetFormatGeometry(s.target, presentation.GeometryPatch{X: &x, Y: &y})
	case "resize":
		w, h := int64(3657600), int64(1828800)
		return s.session.SetFormatGeometry(s.target, presentation.GeometryPatch{Width: &w, Height: &h})
	case "transform":
		angle := int64(5400000)
		yes, no := true, false
		return s.session.SetFormatGeometry(s.target, presentation.GeometryPatch{Rotation: &angle, FlipH: &yes, FlipV: &no})
	case "geometry-refusal":
		x, w := int64(914400), int64(-1)
		s.held = append([]byte{}, s.source...)
		s.refusal = s.session.SetFormatGeometry(s.target, presentation.GeometryPatch{X: &x, Width: &w})
		return nil
	case "run-bold":
		yes := true
		return s.session.SetFormatRun(s.target, 0, 1, presentation.RunFormatPatch{Bold: &yes})
	case "run-unbold":
		no := false
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Bold: &no})
	case "run-italic":
		yes := true
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Italic: &yes})
	case "run-underline":
		v := "none"
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Underline: &v})
	case "run-size":
		v := int64(2400)
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Size: &v})
	case "run-font":
		v := "Aptos Display"
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Latin: &v})
	case "run-color":
		v := "A1B2C3"
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Color: &v})
	case "run-inherit":
		return s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{RemoveDirect: true})
	case "paragraph-align":
		v := "ctr"
		return s.session.SetFormatParagraph(s.target, 0, presentation.ParagraphFormatPatch{Alignment: &v})
	case "paragraph-indent":
		l, i := int64(457200), int64(-228600)
		return s.session.SetFormatParagraph(s.target, 0, presentation.ParagraphFormatPatch{MarginLeft: &l, Indent: &i})
	case "paragraph-spacing":
		b, a := int64(1200), int64(600)
		return s.session.SetFormatParagraph(s.target, 0, presentation.ParagraphFormatPatch{SpaceBefore: &b, SpaceAfter: &a})
	case "paragraph-line":
		v := int64(150000)
		return s.session.SetFormatParagraph(s.target, 0, presentation.ParagraphFormatPatch{LinePercent: &v})
	}
	return fmt.Errorf("unsupported formatting kind %s", s.record.Kind)
}
func (s *formattingState) save() error {
	if s.session == nil {
		return fmt.Errorf("formatting session absent")
	}
	dest, e := os.MkdirTemp("", "pptx-formatting-*")
	if e != nil {
		return e
	}
	defer os.RemoveAll(dest)
	path := filepath.Join(dest, "output.pptx")
	if _, e = s.session.SaveAs(path); e != nil {
		return e
	}
	s.output, e = os.ReadFile(path)
	if e != nil {
		return e
	}
	s.after, e = pptxMembers(s.output)
	if e != nil {
		return e
	}
	_, _, _, _, e = pptxSlideTitles(s.output)
	if e != nil {
		return e
	}
	p, e := presentation.OpenEditing(s.output, packaging.Limits{})
	if e != nil {
		return e
	}
	_, e = p.SlideVisible("ppt/slides/slide2.xml")
	return e
}
func (s *formattingState) checkRefusal() error {
	if s.refusal == nil || len(s.held) == 0 {
		return fmt.Errorf("formatting refusal absent")
	}
	want, e := formatJSON(s.record.Expected)
	if e != nil {
		return e
	}
	var r *packaging.Refusal
	if !errors.As(s.refusal, &r) || r.Kind != want["refusal"] {
		return fmt.Errorf("refusal=%v want %v", s.refusal, want["refusal"])
	}
	if s.target == nil {
		return fmt.Errorf("held target absent")
	}
	if len(s.held) == 0 || !bytes.Equal(s.held, s.source) {
		return fmt.Errorf("held source changed")
	}
	// The original run has direct b=1. Exercise the same held target with a
	// same-value edit after refusal; it must remain usable and byte-identical.
	bold := true
	if e := s.session.SetFormatRun(s.target, 0, 0, presentation.RunFormatPatch{Bold: &bold}); e != nil {
		return fmt.Errorf("held format target unusable: %w", e)
	}
	return nil
}
func (s *formattingState) expected(want map[string]any) error {
	if s.after == nil {
		return fmt.Errorf("saved archive absent")
	}
	if _, refusal := want["refusal"]; refusal {
		return s.checkRefusal()
	}
	part := s.after["ppt/slides/slide2.xml"]
	doc, e := losslessxml.Parse(part)
	if e != nil {
		return e
	}
	root := doc.Elements()[0]
	if strings.HasPrefix(s.record.Kind, "hide") || s.record.Kind == "unhide" || strings.HasPrefix(s.record.Kind, "visibility-") {
		if show, ok := want["show"].(string); ok && pptxAttr(root, "", "show") != show {
			return fmt.Errorf("show=%q want %q", pptxAttr(root, "", "show"), show)
		}
		if hidden, ok := want["hidden"].(bool); ok {
			if actual := pptxAttr(root, "", "show") == "0"; actual != hidden {
				return fmt.Errorf("hidden=%v want %v", actual, hidden)
			}
		}
		if s.record.Kind == "visibility-order" {
			paths, _, _, titles, e := pptxSlideTitles(s.output)
			if e != nil {
				return e
			}
			if !reflect.DeepEqual(titles, []string{"Gamma", "Alpha", "Beta"}) || len(paths) != 3 || paths[2] != "ppt/slides/slide2.xml" || pptxAttr(root, "", "show") != "0" {
				return fmt.Errorf("hidden identity/order %v %v", paths, titles)
			}
		}
		return nil
	}
	shape, e := pptxShape(doc, 4)
	if e != nil {
		return e
	}
	var selected losslessxml.Element
	switch {
	case strings.HasPrefix(s.record.Kind, "run-"):
		body := pptxChild(doc, shape, pptxElem(pptxP, "txBody"))
		if len(body) != 1 {
			return fmt.Errorf("text body")
		}
		ps := pptxChild(doc, body[0], pptxElem(pptxA, "p"))
		if len(ps) != 2 {
			return fmt.Errorf("paragraph count")
		}
		rs := pptxChild(doc, ps[0], pptxElem(pptxA, "r"))
		if len(rs) != 2 {
			return fmt.Errorf("run count")
		}
		idx := 0
		if s.record.Kind == "run-bold" {
			idx = 1
		}
		props := pptxChild(doc, rs[idx], pptxElem(pptxA, "rPr"))
		if len(props) != 1 {
			return fmt.Errorf("run property count")
		}
		selected = props[0]
	case strings.HasPrefix(s.record.Kind, "paragraph-"):
		body := pptxChild(doc, shape, pptxElem(pptxP, "txBody"))
		if len(body) != 1 {
			return fmt.Errorf("text body")
		}
		ps := pptxChild(doc, body[0], pptxElem(pptxA, "p"))
		if len(ps) != 2 {
			return fmt.Errorf("paragraph count")
		}
		props := pptxChild(doc, ps[0], pptxElem(pptxA, "pPr"))
		if len(props) != 1 {
			return fmt.Errorf("paragraph property count")
		}
		selected = props[0]
	case strings.HasPrefix(s.record.Kind, "move") || s.record.Kind == "resize" || s.record.Kind == "transform":
		sppr := pptxChild(doc, shape, pptxElem(pptxP, "spPr"))
		if len(sppr) != 1 {
			return fmt.Errorf("shape property")
		}
		xs := pptxChild(doc, sppr[0], pptxElem(pptxA, "xfrm"))
		if len(xs) != 1 {
			return fmt.Errorf("transform count")
		}
		selected = xs[0]
	}
	if selected.Ordinal() < 0 {
		return fmt.Errorf("selected property absent")
	}
	if e := formattingChildOrder(doc, selected, s.record.Kind); e != nil {
		return e
	}
	for key, value := range want {
		if key == "absent" || key == "retainedLanguage" || key == "before" || key == "after" || key == "linePercent" || key == "typeface" || key == "color" || key == "x" || key == "y" || key == "width" || key == "height" {
			continue
		}
		if key == "rotation" {
			key = "rot"
		}
		if key == "flipH" || key == "flipV" {
			v := "0"
			if value == true {
				v = "1"
			}
			if pptxAttr(selected, "", key) != v {
				return fmt.Errorf("direct %s=%q want %s", key, pptxAttr(selected, "", key), v)
			}
			continue
		}
		if formatScalar(value) != pptxAttr(selected, "", key) {
			return fmt.Errorf("direct %s=%q want %v", key, pptxAttr(selected, "", key), value)
		}
	}
	if _, ok := want["x"]; ok {
		off := pptxChild(doc, selected, pptxElem(pptxA, "off"))
		ext := pptxChild(doc, selected, pptxElem(pptxA, "ext"))
		if len(off) != 1 || len(ext) != 1 {
			return fmt.Errorf("transform geometry children")
		}
		for _, item := range []struct {
			label string
			node  losslessxml.Element
			attr  string
		}{{"x", off[0], "x"}, {"y", off[0], "y"}, {"width", ext[0], "cx"}, {"height", ext[0], "cy"}} {
			if v, ok := want[item.label]; ok && pptxAttr(item.node, "", item.attr) != formatScalar(v) {
				return fmt.Errorf("geometry %s=%q want %v", item.label, pptxAttr(item.node, "", item.attr), v)
			}
		}
	}
	if v, ok := want["typeface"].(string); ok {
		latin := pptxChild(doc, selected, pptxElem(pptxA, "latin"))
		if len(latin) != 1 || pptxAttr(latin[0], "", "typeface") != v {
			return fmt.Errorf("direct Latin font")
		}
	}
	if v, ok := want["color"].(string); ok {
		fill := pptxChild(doc, selected, pptxElem(pptxA, "solidFill"))
		if len(fill) != 1 {
			return fmt.Errorf("direct solid fill")
		}
		srgb := pptxChild(doc, fill[0], pptxElem(pptxA, "srgbClr"))
		if len(srgb) != 1 || pptxAttr(srgb[0], "", "val") != v {
			return fmt.Errorf("direct sRGB color")
		}
	}
	for _, item := range []struct{ key, container, leaf string }{{"before", "spcBef", "spcPts"}, {"after", "spcAft", "spcPts"}, {"linePercent", "lnSpc", "spcPct"}} {
		if v, ok := want[item.key]; ok {
			parent := pptxChild(doc, selected, pptxElem(pptxA, item.container))
			if len(parent) != 1 {
				return fmt.Errorf("%s container", item.key)
			}
			leaf := pptxChild(doc, parent[0], pptxElem(pptxA, item.leaf))
			if len(leaf) != 1 || pptxAttr(leaf[0], "", "val") != formatScalar(v) {
				return fmt.Errorf("%s=%v want %v", item.key, leaf, v)
			}
		}
	}
	if lang, ok := want["retainedLanguage"].(string); ok && pptxAttr(selected, "", "lang") != lang {
		return fmt.Errorf("language drift")
	}
	if absent, ok := want["absent"].([]any); ok {
		for _, item := range absent {
			key := item.(string)
			if key == "latin" || key == "solidFill" {
				if len(pptxChild(doc, selected, pptxElem(pptxA, key))) != 0 {
					return fmt.Errorf("direct %s remained", key)
				}
			} else if pptxAttr(selected, "", key) != "" {
				return fmt.Errorf("direct %s remained", key)
			}
		}
	}
	return nil
}
func (s *formattingState) custody() error {
	if s.after == nil {
		return fmt.Errorf("saved archive absent")
	}
	allowed := map[string]bool{}
	for _, n := range s.record.ChangedMembers {
		allowed[n] = true
	}
	actual := map[string]bool{}
	if len(s.after) != len(s.before) {
		return fmt.Errorf("member name set changed")
	}
	for n, b := range s.before {
		v, ok := s.after[n]
		if !ok {
			return fmt.Errorf("member removed %s", n)
		}
		if !bytes.Equal(b, v) {
			actual[n] = true
			if !allowed[n] {
				return fmt.Errorf("unrelated member changed %s", n)
			}
		}
	}
	if !reflect.DeepEqual(actual, allowed) {
		return fmt.Errorf("changed member set %v want %v", actual, allowed)
	}
	beforeEdges, e := pptxEdgeMap(s.source)
	if e != nil {
		return e
	}
	afterEdges, e := pptxEdgeMap(s.output)
	if e != nil {
		return e
	}
	if !reflect.DeepEqual(beforeEdges, afterEdges) {
		return fmt.Errorf("package graph changed")
	}
	oldPaths, oldRIDs, oldIDs, _, e := pptxSlideTitles(s.source)
	if e != nil {
		return e
	}
	newPaths, newRIDs, newIDs, _, e := pptxSlideTitles(s.output)
	if e != nil {
		return e
	}
	if len(oldPaths) != len(newPaths) {
		return fmt.Errorf("slide count drift")
	}
	if s.record.Kind == "visibility-order" {
		for i, j := range []int{2, 0, 1} {
			if newPaths[i] != oldPaths[j] || newRIDs[i] != oldRIDs[j] || newIDs[i] != oldIDs[j] {
				return fmt.Errorf("slide identity drift at %d", i)
			}
		}
	} else if !reflect.DeepEqual(oldPaths, newPaths) || !reflect.DeepEqual(oldRIDs, newRIDs) || !reflect.DeepEqual(oldIDs, newIDs) {
		return fmt.Errorf("slide identity drift")
	}
	if !bytes.Equal(s.before["ppt/slides/slide2.xml"], s.after["ppt/slides/slide2.xml"]) {
		if e = formattingMaskedSlideSame(s.before["ppt/slides/slide2.xml"], s.after["ppt/slides/slide2.xml"], s.record.Kind); e != nil {
			return e
		}
	}
	if s.record.Kind == "visibility-order" {
		if e = formattingSlideListSame(s.before["ppt/presentation.xml"], s.after["ppt/presentation.xml"]); e != nil {
			return e
		}
	}
	if !bytes.Equal(s.caller, s.callerBefore) || !bytes.Equal(s.source, s.callerBefore) {
		return fmt.Errorf("caller archive mutated")
	}
	if !bytes.Equal(s.source, s.held) && len(s.held) > 0 {
		return fmt.Errorf("held source changed")
	}
	return nil
}
