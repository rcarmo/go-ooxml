package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type creationCaseWorld struct {
	t               *testing.T
	creator         *presentation.ContractPresentation
	baselineCreator *presentation.ContractPresentation
	input           *contract20TitleWorld
	anchor          *presentation.ContractTitleAnchor
	source          []byte
	original        map[string][]byte
	prior           []byte
	priorPath       string
	savedPath       string
	saved           map[string][]byte
	graph           packaging.Graph
	appendOnly      map[string][]byte
	invalid, unsafe error
}

func (w *creationCaseWorld) archive(path string) error {
	b, e := os.ReadFile(path)
	if e != nil {
		return e
	}
	w.saved, e = blankMemberPayloads(b)
	if e != nil {
		return e
	}
	w.graph, e = contract20TableGraph(b)
	return e
}
func (w *creationCaseWorld) openRecipe(id string) error {
	w.input = &contract20TitleWorld{temp: w.t.TempDir()}
	if e := w.input.input(id); e != nil {
		return e
	}
	w.source = bytes.Clone(w.input.source)
	w.original = w.input.before
	return nil
}
func (w *creationCaseWorld) saveFaults(save func(string) error) error {
	if w.savedPath == "" {
		return fmt.Errorf("save fault baseline not recorded")
	}
	baselineBytes, e := os.ReadFile(w.savedPath)
	if e != nil {
		return e
	}
	baselineMembers, e := blankMemberPayloads(baselineBytes)
	if e != nil {
		return e
	}
	baselineGraph, e := contract20TableGraph(baselineBytes)
	if e != nil {
		return e
	}
	dir := w.t.TempDir()
	absent := filepath.Join(dir, "absent-dir", "result.pptx")
	if e := save(absent); e == nil {
		return fmt.Errorf("actual absent-path save did not fail")
	}
	if _, e := os.Stat(absent); !os.IsNotExist(e) {
		return fmt.Errorf("failed save published absent path: %v", e)
	}
	old := []byte("previous destination bytes")
	target := filepath.Join(dir, "prior.pptx")
	if e := os.WriteFile(target, old, 0600); e != nil {
		return e
	}
	link := filepath.Join(dir, "existing.pptx")
	if e := os.Symlink(target, link); e != nil {
		return e
	}
	if e := save(link); e == nil {
		return fmt.Errorf("actual existing-path save did not fail")
	}
	got, e := os.ReadFile(target)
	if e != nil || !bytes.Equal(got, old) {
		return fmt.Errorf("previous destination changed: %v", e)
	}
	info, e := os.Lstat(link)
	if e != nil || info.Mode()&os.ModeSymlink == 0 {
		return fmt.Errorf("existing link replaced: %v", e)
	}
	entries, e := os.ReadDir(dir)
	if e != nil || len(entries) != 2 {
		return fmt.Errorf("staging file survived: %v %v", entries, e)
	}
	// Exercise the SAME live session after both actual delivery failures.
	// Reopen and compare full payload inventories and the resolved OPC graph;
	// archive-byte equality is not used as a generated-output gate.
	recovered := filepath.Join(w.t.TempDir(), "post-fault.pptx")
	if e := save(recovered); e != nil {
		return fmt.Errorf("live post-fault save: %w", e)
	}
	got, e = os.ReadFile(recovered)
	if e != nil {
		return e
	}
	members, e := blankMemberPayloads(got)
	if e != nil {
		return e
	}
	graph, e := contract20TableGraph(got)
	if e != nil {
		return e
	}
	if !reflect.DeepEqual(members, baselineMembers) || !reflect.DeepEqual(graph, baselineGraph) {
		return fmt.Errorf("live session differs after failed save")
	}
	if w.input != nil && w.anchor != nil {
		order, e := creationSlideOrder(baselineMembers, baselineGraph)
		if e != nil {
			return e
		}
		anchor, e := w.input.session.FindContractTitle(order[0])
		if e != nil {
			return fmt.Errorf("post-fault fresh anchor: %w", e)
		}
		if anchor.Text() == "" {
			return fmt.Errorf("post-fault empty title anchor")
		}
		changed, e := w.input.session.ReplaceContractTitleSpan(anchor, 0, len([]rune(anchor.Text())), anchor.Text())
		if e != nil || changed {
			return fmt.Errorf("post-fault fresh anchor no-op: changed=%v error=%v", changed, e)
		}
	}
	return nil
}
func creationCheckRefusal(err error, kind string) error {
	var r *packaging.Refusal
	if !errors.As(err, &r) || r.Kind != kind {
		return fmt.Errorf("expected %s typed refusal, got %v", kind, err)
	}
	return nil
}
func creationPlaceholderOperands(data []byte, title, subtitle string) error {
	d, err := losslessxml.Parse(data)
	if err != nil {
		return err
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	values := map[string][]string{}
	for _, ph := range d.Elements() {
		if ph.Name() != p("ph") {
			continue
		}
		kind := ""
		for _, attr := range ph.Attributes() {
			if attr.Name == (xml.Name{Local: "type"}) {
				kind = attr.Value
			}
		}
		if kind != "title" && kind != "ctrTitle" && kind != "subTitle" {
			continue
		}
		var shape losslessxml.Element
		found := false
		for ancestor, ok := ph.Parent(); ok; ancestor, ok = ancestor.Parent() {
			if ancestor.Name() == p("sp") {
				shape = ancestor
				found = true
				break
			}
		}
		if !found {
			return fmt.Errorf("placeholder without shape")
		}
		for _, leaf := range d.Elements() {
			if leaf.Name() != a("t") {
				continue
			}
			v, ok := leaf.Text()
			if !ok {
				continue
			}
			for ancestor, ok := leaf.Parent(); ok; ancestor, ok = ancestor.Parent() {
				if ancestor == shape {
					values[kind] = append(values[kind], v)
					break
				}
			}
		}
	}
	titles := append(values["ctrTitle"], values["title"]...)
	if !reflect.DeepEqual(titles, []string{title}) || !reflect.DeepEqual(values["subTitle"], []string{subtitle}) {
		return fmt.Errorf("placeholder title/subtitle operands %v", values)
	}
	return nil
}
func creationReopened(path string, want [][2]string) error {
	archive, e := os.ReadFile(path)
	if e != nil {
		return e
	}
	members, e := blankMemberPayloads(archive)
	if e != nil {
		return e
	}
	graph, e := contract20TableGraph(archive)
	if e != nil {
		return e
	}
	order, e := creationSlideOrder(members, graph)
	if e != nil {
		return e
	}
	if len(order) != len(want) {
		return fmt.Errorf("reopened resolved order %d, want %d", len(order), len(want))
	}
	for i, pair := range want {
		if e := creationPlaceholderOperands(members[order[i]], pair[0], pair[1]); e != nil {
			return e
		}
	}
	deck, e := presentation.Open(path)
	if e != nil {
		return e
	}
	defer deck.Close()
	if deck.SlideCount() != len(want) {
		return fmt.Errorf("reopen slides %d, want %d", deck.SlideCount(), len(want))
	}
	for i, pair := range want {
		slide, e := deck.Slide(i + 1)
		if e != nil {
			return e
		}
		texts := []string{}
		for _, shape := range slide.Shapes() {
			texts = append(texts, shape.Text())
		}
		if !reflect.DeepEqual(texts, []string{pair[0], pair[1]}) {
			return fmt.Errorf("slide %d text %q", i, texts)
		}
	}
	return nil
}
func creationGraphDelta(before, after packaging.Graph, oldSlide, newSlide string) error {
	beforeParts := map[string]string{}
	for _, part := range before.Parts {
		beforeParts[part.Name] = part.ContentType
	}
	additions := map[string]string{}
	for _, part := range after.Parts {
		if ct, ok := beforeParts[part.Name]; ok {
			if ct != part.ContentType {
				return fmt.Errorf("old MIME changed %s", part.Name)
			}
			delete(beforeParts, part.Name)
		} else {
			additions[part.Name] = part.ContentType
		}
	}
	if len(beforeParts) != 0 || len(additions) != 2 || additions[newSlide] != packaging.ContentTypeSlide || additions[packaging.RelationshipsPathForPart(newSlide)] != "application/vnd.openxmlformats-package.relationships+xml" {
		return fmt.Errorf("member types removed=%v added=%v", beforeParts, additions)
	}
	old := map[string]packaging.Edge{}
	for _, edge := range before.Edges {
		old[edge.Source+"|"+edge.ID] = edge
	}
	added := []packaging.Edge{}
	for _, edge := range after.Edges {
		key := edge.Source + "|" + edge.ID
		if prior, ok := old[key]; ok {
			if prior != edge {
				return fmt.Errorf("old edge changed %s", key)
			}
			delete(old, key)
		} else {
			added = append(added, edge)
		}
	}
	if len(old) != 0 || len(added) != 2 {
		return fmt.Errorf("edges removed=%v added=%v", old, added)
	}
	layout := ""
	for _, edge := range before.Edges {
		if edge.Source == oldSlide && edge.Type == packaging.RelTypeSlideLayout {
			layout = edge.ResolvedPart
		}
	}
	slides, layouts := 0, 0
	for _, edge := range added {
		if edge.External {
			return fmt.Errorf("external added edge")
		}
		switch {
		case edge.Source == "ppt/presentation.xml" && edge.Type == packaging.RelTypeSlide && edge.ResolvedPart == newSlide:
			slides++
		case edge.Source == newSlide && edge.Type == packaging.RelTypeSlideLayout && edge.ResolvedPart == layout:
			layouts++
		default:
			return fmt.Errorf("unexpected new edge %+v", edge)
		}
	}
	if slides != 1 || layouts != 1 {
		return fmt.Errorf("missing append edges %d/%d", slides, layouts)
	}
	return nil
}
func (w *creationCaseWorld) selectedSteps(id string) map[string]func() error {
	pairs := [][2]string{{"Created title 1", "Created subtitle 1"}, {"Created title 2", "Created subtitle 2"}}
	expected := [][2]string{{"Edited original title", "Original subtitle"}, {"Appended title", "Appended subtitle"}}
	if id == "@id-pptx-create-new-minimal" {
		return map[string]func() error{
			"no input package is supplied to the production presentation creator": func() error { var e error; w.creator, e = presentation.NewContractPresentation(); return e },
			`the production creator authors exactly two text slides with JSON [["Created title 1","Created subtitle 1"],["Created title 2","Created subtitle 2"]]`: func() error {
				for _, pair := range pairs {
					if e := w.creator.AddTitleSlide(pair[0], pair[1]); e != nil {
						return e
					}
				}
				if w.creator.SlideCount() != 2 {
					return fmt.Errorf("creation count")
				}
				return nil
			},
			"the resulting presentation is saved to a new path and independently parsed and reopened": func() error {
				w.savedPath = filepath.Join(w.t.TempDir(), "created.pptx")
				if e := w.creator.SaveAs(w.savedPath); e != nil {
					return e
				}
				if e := w.archive(w.savedPath); e != nil {
					return e
				}
				return creationReopened(w.savedPath, pairs)
			},
			"the reopened slide count is 2 with exactly one master and one title-layout definition": func() error {
				counts := map[string]int{}
				for _, part := range w.graph.Parts {
					counts[part.ContentType]++
				}
				if counts[packaging.ContentTypeSlide] != 2 || counts[packaging.ContentTypeSlideMaster] != 1 || counts[packaging.ContentTypeSlideLayout] != 1 {
					return fmt.Errorf("created roles %v", counts)
				}
				return creationReopened(w.savedPath, pairs)
			},
			`the exact ordered title and subtitle placeholder text equals JSON [["Created title 1","Created subtitle 1"],["Created title 2","Created subtitle 2"]]`: func() error {
				order, err := creationSlideOrder(w.saved, w.graph)
				if err != nil {
					return err
				}
				if len(order) != len(pairs) {
					return fmt.Errorf("created slide order count %d", len(order))
				}
				for i, pair := range pairs {
					if err := creationPlaceholderOperands(w.saved[order[i]], pair[0], pair[1]); err != nil {
						return err
					}
				}
				return nil
			},
			"the support graph follows contract20-creation-policy.json: exactly one presentation master title-layout and theme; optional presentation-properties view-properties table-style-definitions core-properties and application-properties have at most one part each, and no other roles or unknown edges exist": func() error { return contract20CreationPolicy(w.saved, w.graph, 2) },
			"each slide owns exactly one resolved layout edge and the master and layout link to each other without duplicate IDs": func() error {
				counts := map[string]int{}
				for _, edge := range w.graph.Edges {
					if edge.External || edge.ResolvedPart == "" {
						return fmt.Errorf("unresolved created relationship")
					}
					key := edge.Source + "|" + edge.ID
					counts[key]++
					if counts[key] != 1 {
						return fmt.Errorf("duplicate ID %s", key)
					}
				}
				order, err := creationSlideOrder(w.saved, w.graph)
				if err != nil {
					return err
				}
				if len(order) != 2 {
					return fmt.Errorf("created order length %d", len(order))
				}
				for i, source := range order {
					n := 0
					for _, edge := range w.graph.Edges {
						if edge.Source == source && edge.Type == packaging.RelTypeSlideLayout {
							n++
						}
					}
					if n != 1 {
						return fmt.Errorf("slide %d layout edges %d", i, n)
					}
				}
				return contract20CreationPolicy(w.saved, w.graph, 2)
			},
			"no unrequested notes media external links tables or extra slides are authored": func() error {
				for part, data := range w.saved {
					if strings.Contains(part, "notes") || strings.Contains(part, "media") || bytes.Contains(data, []byte("<a:tbl>")) {
						return fmt.Errorf("unrequested part %s", part)
					}
				}
				for _, edge := range w.graph.Edges {
					if edge.External {
						return fmt.Errorf("external edge")
					}
				}
				return nil
			},
			"a failed save leaves an existing destination unchanged or a previously absent destination absent": func() error {
				if e := w.saveFaults(w.creator.SaveAs); e != nil {
					return e
				}
				return creationReopened(w.savedPath, pairs)
			},
		}
	}
	if id == "@id-pptx-create-anchor-preservation" {
		return map[string]func() error{
			`the contract20 derived PPTX recipe "title-subtitle-source" is created in memory from its sealed fixture`: func() error { return w.openRecipe("title-subtitle-source") },
			`the first title anchor is held while its text is JSON "Original title" and original member identities are recorded`: func() error {
				var e error
				order, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				if len(order) != 1 {
					return fmt.Errorf("source order count %d", len(order))
				}
				w.anchor, e = w.input.session.FindContractTitle(order[0])
				if e != nil {
					return e
				}
				if w.anchor.Text() != "Original title" {
					return fmt.Errorf("wrong held title")
				}
				return nil
			},
			`the production presentation editor appends title JSON "Appended title" and subtitle JSON "Appended subtitle"`: func() error { return w.input.session.AppendContractTitleSlide("Appended title", "Appended subtitle") },
			"the original slide part and relationships remain literal and the held original title anchor remains valid": func() error {
				path := filepath.Join(w.t.TempDir(), "append.pptx")
				if _, e := w.input.session.SaveAs(path); e != nil {
					return e
				}
				b, e := os.ReadFile(path)
				if e != nil {
					return e
				}
				w.appendOnly, e = blankMemberPayloads(b)
				if e != nil {
					return e
				}
				oldOrder, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				if len(oldOrder) != 1 {
					return fmt.Errorf("source order count %d", len(oldOrder))
				}
				for _, name := range []string{oldOrder[0], packaging.RelationshipsPathForPart(oldOrder[0])} {
					if !bytes.Equal(w.appendOnly[name], w.original[name]) {
						return fmt.Errorf("old slide part changed: %s", name)
					}
				}
				if w.anchor.Text() != "Original title" {
					return fmt.Errorf("held anchor invalidated")
				}
				return nil
			},
			`the production anchored editor uses that held anchor to replace JSON "Original title" with JSON "Edited original title"`: func() error {
				changed, e := w.input.session.ReplaceContractTitleSpan(w.anchor, 0, len("Original title"), "Edited original title")
				if e != nil {
					return e
				}
				if !changed {
					return fmt.Errorf("held replacement no-op")
				}
				return nil
			},
			"the result is saved to a distinct new path and independently parsed and reopened": func() error {
				w.savedPath = filepath.Join(w.t.TempDir(), "result.pptx")
				if _, e := w.input.session.SaveAs(w.savedPath); e != nil {
					return e
				}
				if e := w.archive(w.savedPath); e != nil {
					return e
				}
				return creationReopened(w.savedPath, expected)
			},
			"the source fixture, caller archive and operation operands remain unchanged": func() error {
				if !bytes.Equal(w.input.original, w.input.source) || !bytes.Equal(w.source, w.input.source) {
					return fmt.Errorf("source archive changed")
				}
				base, e := blankMemberPayloads(w.input.baseFixture)
				if e != nil {
					return e
				}
				if !reflect.DeepEqual(base, w.original) {
					return fmt.Errorf("fixture mutated")
				}
				if w.anchor.Text() != "Original title" {
					return fmt.Errorf("operand changed")
				}
				return nil
			},
			"only the sealed original member and lexical span allowances differ; all unrelated member payloads remain literal": func() error {
				oldOrder, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				if len(oldOrder) != 1 {
					return fmt.Errorf("source order count %d", len(oldOrder))
				}
				allow := map[string]bool{oldOrder[0]: true, "ppt/presentation.xml": true, "ppt/_rels/presentation.xml.rels": true, "[Content_Types].xml": true}
				for name, data := range w.original {
					after, ok := w.saved[name]
					if !ok || !allow[name] && !bytes.Equal(data, after) {
						return fmt.Errorf("unrelated member %s", name)
					}
				}
				return nil
			},
			"all saved OPC relationships and content types resolve with original identities preserved": func() error {
				oldOrder, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				newOrder, e := creationSlideOrder(w.saved, w.graph)
				if e != nil {
					return e
				}
				if len(oldOrder) != 1 || len(newOrder) != 2 {
					return fmt.Errorf("append order count")
				}
				return creationGraphDelta(w.input.graph, w.graph, oldOrder[0], newOrder[1])
			},
			"refusals and save faults publish no partial destination and leave the session and held unaffected targets usable": func() error {
				if err := creationCheckRefusal(w.input.session.AppendContractTitleSlide("bad\x00", "Bad subtitle"), "PPTX_ARGUMENT_INVALID"); err != nil {
					return err
				}
				save := func(path string) error { _, e := w.input.session.SaveAs(path); return e }
				if e := w.saveFaults(save); e != nil {
					return e
				}
				order, e := creationSlideOrder(w.saved, w.graph)
				if e != nil {
					return e
				}
				if len(order) != 2 {
					return fmt.Errorf("append order count")
				}
				if _, e = w.input.session.FindContractTitle(order[1]); e != nil {
					return e
				}
				return creationReopened(w.savedPath, expected)
			},
			`the exact reopened placeholder order equals JSON [["Edited original title","Original subtitle"],["Appended title","Appended subtitle"]]`: func() error {
				if err := creationReopened(w.savedPath, expected); err != nil {
					return err
				}
				order, err := creationSlideOrder(w.saved, w.graph)
				if err != nil {
					return err
				}
				if len(order) != len(expected) {
					return fmt.Errorf("append order length %d", len(order))
				}
				for i, pair := range expected {
					if err := creationPlaceholderOperands(w.saved[order[i]], pair[0], pair[1]); err != nil {
						return err
					}
				}
				return nil
			},
			"the old slide retains its original IDs relationship IDs targets layout links and payload outside the selected title run": func() error {
				oldOrder, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				if len(oldOrder) != 1 {
					return fmt.Errorf("source order count")
				}
				oldName := oldOrder[0]
				oldRels := packaging.RelationshipsPathForPart(oldName)
				old := w.original[oldName]
				if bytes.Count(old, []byte("Original title")) != 1 {
					return fmt.Errorf("source title not unique")
				}
				want := bytes.Replace(old, []byte("Original title"), []byte("Edited original title"), 1)
				if !bytes.Equal(w.saved[oldName], want) || !bytes.Equal(w.saved[oldRels], w.original[oldRels]) {
					return fmt.Errorf("old slide lexical custody")
				}
				return nil
			},
			"exactly one new slide and its one layout relationship part are added with unique valid identities; no member is removed": func() error {
				oldOrder, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				newOrder, e := creationSlideOrder(w.saved, w.graph)
				if e != nil {
					return e
				}
				if len(oldOrder) != 1 || len(newOrder) != 2 || len(w.saved) != len(w.original)+2 {
					return fmt.Errorf("append order/member count")
				}
				newSlide := newOrder[1]
				newRels := packaging.RelationshipsPathForPart(newSlide)
				for name := range w.saved {
					if _, ok := w.original[name]; !ok && name != newSlide && name != newRels {
						return fmt.Errorf("unexpected added member %s", name)
					}
				}
				return creationGraphDelta(w.input.graph, w.graph, oldOrder[0], newSlide)
			},
			"only original presentation list relationships content types and selected first-slide text spans may change": func() error {
				oldOrder, e := creationSlideOrder(w.original, w.input.graph)
				if e != nil {
					return e
				}
				newOrder, e := creationSlideOrder(w.saved, w.graph)
				if e != nil {
					return e
				}
				if len(oldOrder) != 1 || len(newOrder) != 2 {
					return fmt.Errorf("append order count")
				}
				return creationAppendInsertions(w.original, w.saved, w.input.graph, w.graph, oldOrder[0], newOrder[1])
			},
		}
	}
	return map[string]func() error{
		`the production creator has a new empty session and a second session opens contract20 derived PPTX recipe "unsafe-title-layout"`: func() error {
			var e error
			w.creator, e = presentation.NewContractPresentation()
			if e != nil {
				return e
			}
			w.baselineCreator, e = presentation.NewContractPresentation()
			if e != nil {
				return e
			}
			return w.openRecipe("unsafe-title-layout")
		},
		"both source session states and prior destination bytes are recorded": func() error {
			w.prior = []byte("prior destination")
			w.priorPath = filepath.Join(w.t.TempDir(), "prior.pptx")
			if e := os.WriteFile(w.priorPath, w.prior, 0600); e != nil {
				return e
			}
			return nil
		},
		`the new session attempts title JSON "bad\u0000" and subtitle JSON "Bad subtitle"`: func() error { w.invalid = w.creator.AddTitleSlide("bad\x00", "Bad subtitle"); return nil },
		`that operation returns exactly typed refusal category "PPTX_ARGUMENT_INVALID" with no new slide or partial result`: func() error {
			if e := creationCheckRefusal(w.invalid, "PPTX_ARGUMENT_INVALID"); e != nil {
				return e
			}
			if w.creator.SlideCount() != 0 || w.baselineCreator.SlideCount() != 0 {
				return fmt.Errorf("creator empty-state observable inventory changed")
			}
			for _, creator := range []*presentation.ContractPresentation{w.creator, w.baselineCreator} {
				path := filepath.Join(w.t.TempDir(), "empty.pptx")
				if err := creationCheckRefusal(creator.SaveAs(path), "PPTX_CREATION_UNSUPPORTED"); err != nil {
					return err
				}
				if _, err := os.Stat(path); !os.IsNotExist(err) {
					return fmt.Errorf("empty creator published output: %v", err)
				}
			}
			return nil
		},
		`the unsafe-layout session attempts append with title JSON "Unsafe next title" and subtitle JSON "Unsafe next subtitle"`: func() error {
			w.unsafe = w.input.session.AppendContractTitleSlide("Unsafe next title", "Unsafe next subtitle")
			return nil
		},
		`that operation returns exactly typed refusal category "PPTX_LAYOUT_UNSAFE" with no changed result`: func() error { return creationCheckRefusal(w.unsafe, "PPTX_LAYOUT_UNSAFE") },
		"both complete session member states and all source relationships remain unchanged": func() error {
			path := filepath.Join(w.t.TempDir(), "unchanged.pptx")
			w.savedPath = path
			if _, e := w.input.session.SaveAs(path); e != nil {
				return e
			}
			if e := w.archive(path); e != nil {
				return e
			}
			if !reflect.DeepEqual(w.saved, w.original) || !reflect.DeepEqual(w.graph, w.input.graph) || !bytes.Equal(w.source, w.input.source) {
				return fmt.Errorf("refusal changed session/source or graph")
			}
			return nil
		},
		"no refusal or save fault creates or replaces a destination and held unaffected reads remain usable": func() error {
			if e := creationCheckRefusal(w.creator.SaveAs(w.priorPath), "PPTX_CREATION_UNSUPPORTED"); e != nil {
				return e
			}
			before, e := os.ReadFile(w.priorPath)
			if e != nil || !bytes.Equal(before, w.prior) {
				return fmt.Errorf("prior destination changed: %v", e)
			}
			save := func(path string) error { _, e := w.input.session.SaveAs(path); return e }
			if e := w.saveFaults(save); e != nil {
				return e
			}
			// Independently observable complete empty-creator comparison: both
			// sessions receive the same first valid operation after the refusal.
			for _, creator := range []*presentation.ContractPresentation{w.creator, w.baselineCreator} {
				if e := creator.AddTitleSlide("Recovered title", "Recovered subtitle"); e != nil {
					return e
				}
			}
			var baseMembers map[string][]byte
			var baseGraph packaging.Graph
			for i, creator := range []*presentation.ContractPresentation{w.baselineCreator, w.creator} {
				path := filepath.Join(w.t.TempDir(), "recovered.pptx")
				if e := creator.SaveAs(path); e != nil {
					return e
				}
				archive, e := os.ReadFile(path)
				if e != nil {
					return e
				}
				members, e := blankMemberPayloads(archive)
				if e != nil {
					return e
				}
				graph, e := contract20TableGraph(archive)
				if e != nil {
					return e
				}
				if e = contract20CreationPolicy(members, graph, 1); e != nil {
					return e
				}
				if i == 0 {
					baseMembers, baseGraph = members, graph
				} else if !reflect.DeepEqual(members, baseMembers) || !reflect.DeepEqual(graph, baseGraph) {
					return fmt.Errorf("refused empty creator differed from fresh baseline")
				}
			}
			order, e := creationSlideOrder(w.original, w.input.graph)
			if e != nil {
				return e
			}
			if len(order) != 1 {
				return fmt.Errorf("refusal source order")
			}
			anchor, e := w.input.session.FindContractTitle(order[0])
			if e != nil || anchor == nil || anchor.Text() != "Original title" {
				return fmt.Errorf("held read after refusal: %v", e)
			}
			return nil
		},
	}
}
func TestContract20CreationCases(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, e := contract20Inventory()
	if e != nil {
		t.Fatal(e)
	}
	for _, id := range []string{"@id-pptx-create-new-minimal", "@id-pptx-create-anchor-preservation", "@id-pptx-create-refusals"} {
		t.Run(strings.TrimPrefix(id, "@id-"), func(t *testing.T) {
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
				t.Fatalf("selected case count for %s = %d", id, count)
			}
			world := &creationCaseWorld{t: t}
			defer func() {
				if world.creator != nil {
					_ = world.creator.Close()
				}
				if world.baselineCreator != nil {
					_ = world.baselineCreator.Close()
				}
			}()
			steps := world.selectedSteps(id)
			ledger, err := readContract20Ledger()
			if err != nil {
				t.Fatal(err)
			}
			bound := 0
			for _, file := range ledger.Files {
				if file.Path != "workflows/pptx/creation.feature" {
					continue
				}
				for _, scenario := range file.Scenarios {
					if scenario.ID != id {
						continue
					}
					for _, row := range scenario.After {
						if row.Name != want.Name {
							continue
						}
						for _, step := range row.Steps {
							if steps[step.Text] == nil {
								t.Fatalf("unbound contract step %q", step.Text)
							}
							bound++
						}
					}
				}
			}
			if len(steps) != want.Steps || bound != want.Steps {
				t.Fatalf("bound steps %d/%d; contract steps %d", len(steps), bound, want.Steps)
			}
			init := func(sc *godog.ScenarioContext) {
				for text, fn := range steps {
					sc.Step("^"+regexp.QuoteMeta(text)+"$", fn)
				}
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) { return ctx, nil })
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-creation", ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: id, Strict: true, Concurrency: 1}}
			if code := suite.Run(); code != 0 {
				t.Errorf("Godog %d: %s", code, output.String())
			}
			var report []reportFeature
			if e := json.Unmarshal(output.Bytes(), &report); e != nil {
				t.Fatalf("report: %v: %s", e, output.String())
			}
			data, e := json.Marshal(report)
			if e != nil {
				t.Fatal(e)
			}
			if e := reconcile(map[caseID]expectedCase{key: want}, data); e != nil {
				t.Error(e)
			}
			if dir := os.Getenv("GO_CONTRACT20_PPTX_REMAINING_RECEIPTS_DIR"); dir != "" && !t.Failed() {
				if e := os.MkdirAll(dir, 0700); e != nil {
					t.Fatal(e)
				}
				path := filepath.Join(dir, strings.TrimPrefix(id, "@id-")+".cucumber.json")
				if e := os.WriteFile(path, data, 0600); e != nil {
					t.Fatal(e)
				}
			}
		})
	}
}
