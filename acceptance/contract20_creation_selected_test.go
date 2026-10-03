package acceptance

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

func contract20CreationSelected(t *testing.T, id string) {
	t.Helper()
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	rows := 0
	for key := range inventory {
		if key.ID == id {
			rows++
		}
	}
	if rows != 1 {
		t.Fatalf("selected creation case %s count %d", id, rows)
	}
}
func contract20CreationRefusal(t *testing.T, err error, kind string) {
	t.Helper()
	var refusal *packaging.Refusal
	if !errors.As(err, &refusal) || refusal.Kind != kind {
		t.Fatalf("want typed %s refusal, got %v", kind, err)
	}
}
func contract20CreationArchive(t *testing.T, path string) ([]byte, map[string][]byte, packaging.Graph) {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	members, err := blankMemberPayloads(b)
	if err != nil {
		t.Fatal(err)
	}
	graph, err := contract20TableGraph(b)
	if err != nil {
		t.Fatal(err)
	}
	return b, members, graph
}
func contract20CreationSaveFaults(t *testing.T, save func(string) error) {
	t.Helper()
	dir := t.TempDir()
	missing := filepath.Join(dir, "missing-parent", "out.pptx")
	if err := save(missing); err == nil {
		t.Fatal("actual failed save unexpectedly succeeded")
	}
	if _, err := os.Stat(missing); !os.IsNotExist(err) {
		t.Fatalf("absent destination published: %v", err)
	}
	// A symlink is an existing destination with independently checkable bytes.
	prior := []byte("pre-existing destination bytes")
	target := filepath.Join(dir, "target.pptx")
	if err := os.WriteFile(target, prior, 0600); err != nil {
		t.Fatal(err)
	}
	link := filepath.Join(dir, "existing.pptx")
	if err := os.Symlink(target, link); err != nil {
		t.Fatal(err)
	}
	if err := save(link); err == nil {
		t.Fatal("existing symlink save unexpectedly succeeded")
	}
	got, err := os.ReadFile(target)
	if err != nil || !bytes.Equal(got, prior) {
		t.Fatalf("existing bytes changed: %v", err)
	}
	if info, err := os.Lstat(link); err != nil || info.Mode()&os.ModeSymlink == 0 {
		t.Fatalf("destination symlink replaced: %v", err)
	}
	entries, err := os.ReadDir(dir)
	if err != nil || len(entries) != 2 {
		t.Fatalf("fault left staging members: %v %v", entries, err)
	}
}
func contract20AppendGraph(t *testing.T, before, after packaging.Graph, oldSlide, newSlide string) {
	t.Helper()
	oldParts := map[string]string{}
	newParts := map[string]string{}
	for _, part := range before.Parts {
		oldParts[part.Name] = part.ContentType
	}
	for _, part := range after.Parts {
		if typ, ok := oldParts[part.Name]; ok {
			if part.ContentType != typ {
				t.Fatalf("old member MIME changed: %s", part.Name)
			}
			delete(oldParts, part.Name)
		} else {
			newParts[part.Name] = part.ContentType
		}
	}
	if len(oldParts) != 0 || len(newParts) != 2 || newParts[newSlide] != packaging.ContentTypeSlide || newParts[packaging.RelationshipsPathForPart(newSlide)] != "application/vnd.openxmlformats-package.relationships+xml" {
		t.Fatalf("part type delta removed=%v new=%v", oldParts, newParts)
	}
	old := map[string]packaging.Edge{}
	added := map[string]packaging.Edge{}
	for _, edge := range before.Edges {
		old[edge.Source+"|"+edge.ID] = edge
	}
	for _, edge := range after.Edges {
		key := edge.Source + "|" + edge.ID
		if prior, ok := old[key]; ok {
			if !reflect.DeepEqual(edge, prior) {
				t.Fatalf("old edge redirected %s", key)
			}
			delete(old, key)
		} else {
			added[key] = edge
		}
	}
	if len(old) != 0 || len(added) != 2 {
		t.Fatalf("edge delta removed=%v added=%v", old, added)
	}
	layout := ""
	for _, edge := range before.Edges {
		if edge.Source == oldSlide && edge.Type == packaging.RelTypeSlideLayout {
			layout = edge.ResolvedPart
		}
	}
	if layout == "" {
		t.Fatal("old layout unresolved")
	}
	seenPresentation, seenLayout := 0, 0
	for _, edge := range added {
		if edge.External {
			t.Fatalf("external added edge %+v", edge)
		}
		switch {
		case edge.Source == "ppt/presentation.xml" && edge.ResolvedPart == newSlide && edge.Type == packaging.RelTypeSlide:
			seenPresentation++
		case edge.Source == newSlide && edge.ResolvedPart == layout && edge.Type == packaging.RelTypeSlideLayout:
			seenLayout++
		default:
			t.Fatalf("unexpected added edge %+v", edge)
		}
	}
	if seenPresentation != 1 || seenLayout != 1 {
		t.Fatalf("missing append edges: %d %d", seenPresentation, seenLayout)
	}
}

// Native alternates harden the writer, but never add selected case credit.
func TestContract20CreationNativeControls(t *testing.T) {
	contract20CreationSelected(t, "@id-pptx-create-anchor-preservation")
	w := &contract20TitleWorld{temp: t.TempDir()}
	if err := w.input("title-subtitle-source"); err != nil {
		t.Fatal(err)
	}
	if err := w.session.AppendContractTitleSlide(" Left & <right> ", " Tail "); err != nil {
		t.Fatal(err)
	}
	dest := filepath.Join(t.TempDir(), "alternate.pptx")
	if _, err := w.session.SaveAs(dest); err != nil {
		t.Fatal(err)
	}
	_, parts, graph := contract20CreationArchive(t, dest)
	order, err := creationSlideOrder(parts, graph)
	if err != nil || len(order) != 2 {
		t.Fatalf("native resolved order %q: %v", order, err)
	}
	for _, text := range []string{" Left &amp; &lt;right&gt; ", " Tail "} {
		if !bytes.Contains(parts[order[1]], []byte(`xml:space="preserve">`+text+`</a:t>`)) {
			t.Fatalf("edge-space or escaped title not retained: %q", text)
		}
	}
	reopened, err := presentation.Open(dest)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	slide, err := reopened.Slide(2)
	if err != nil {
		t.Fatal(err)
	}
	texts := []string{}
	for _, shape := range slide.Shapes() {
		texts = append(texts, shape.Text())
	}
	if !reflect.DeepEqual(texts, []string{" Left & <right> ", " Tail "}) {
		t.Fatalf("alternate reopened %q", texts)
	}
}
func TestContract20CreationSelected(t *testing.T) {
	t.Run("minimal", func(t *testing.T) {
		contract20CreationSelected(t, "@id-pptx-create-new-minimal")
		creator, err := presentation.NewContractPresentation()
		if err != nil {
			t.Fatal(err)
		}
		defer creator.Close()
		pairs := [][2]string{{"Created title 1", "Created subtitle 1"}, {"Created title 2", "Created subtitle 2"}}
		for _, pair := range pairs {
			if err := creator.AddTitleSlide(pair[0], pair[1]); err != nil {
				t.Fatal(err)
			}
		}
		if creator.SlideCount() != 2 {
			t.Fatal("two slides required")
		}
		contract20CreationSaveFaults(t, creator.SaveAs)
		dest := filepath.Join(t.TempDir(), "created.pptx")
		if err := creator.SaveAs(dest); err != nil {
			t.Fatal(err)
		}
		_, members, graph := contract20CreationArchive(t, dest)
		if err := contract20CreationPolicy(members, graph, 2); err != nil {
			t.Fatal(err)
		}
		for part, data := range members {
			if strings.Contains(part, "notes") || strings.Contains(part, "media") || bytes.Contains(data, []byte("<a:tbl>")) {
				t.Fatalf("unrequested part/content %s", part)
			}
		}
		deck, err := presentation.Open(dest)
		if err != nil {
			t.Fatal(err)
		}
		defer deck.Close()
		if deck.SlideCount() != 2 {
			t.Fatal("reopened slide count")
		}
		order, err := creationSlideOrder(members, graph)
		if err != nil || len(order) != len(pairs) {
			t.Fatalf("resolved creation order: %q %v", order, err)
		}
		for i, pair := range pairs {
			slide, e := deck.Slide(i + 1)
			if e != nil {
				t.Fatal(e)
			}
			values := []string{}
			for _, shape := range slide.Shapes() {
				values = append(values, shape.Text())
			}
			if !reflect.DeepEqual(values, []string{pair[0], pair[1]}) {
				t.Fatalf("reopened slide %d text %q", i, values)
			}
			part := members[order[i]]
			if len(part) == 0 || !bytes.Contains(part, []byte(">"+pair[0]+"</a:t>")) || !bytes.Contains(part, []byte(">"+pair[1]+"</a:t>")) {
				t.Fatalf("independent slide %d operands", i)
			}
		}
	})
	t.Run("append-anchor", func(t *testing.T) {
		contract20CreationSelected(t, "@id-pptx-create-anchor-preservation")
		w := &contract20TitleWorld{temp: t.TempDir()}
		if err := w.input("title-subtitle-source"); err != nil {
			t.Fatal(err)
		}
		anchor, err := w.session.FindContractTitle("ppt/slides/slide1.xml")
		if err != nil || anchor.Text() != "Original title" {
			t.Fatalf("original anchor: %v", err)
		}
		if err := w.session.AppendContractTitleSlide("Appended title", "Appended subtitle"); err != nil {
			t.Fatal(err)
		}
		afterAppend := filepath.Join(t.TempDir(), "append-only.pptx")
		if _, err := w.session.SaveAs(afterAppend); err != nil {
			t.Fatal(err)
		}
		_, appendMembers, _ := contract20CreationArchive(t, afterAppend)
		if !bytes.Equal(appendMembers["ppt/slides/slide1.xml"], w.before["ppt/slides/slide1.xml"]) || !bytes.Equal(appendMembers["ppt/slides/_rels/slide1.xml.rels"], w.before["ppt/slides/_rels/slide1.xml.rels"]) {
			t.Fatal("append altered held slide or layout relationship")
		}
		changed, err := w.session.ReplaceContractTitleSpan(anchor, 0, len("Original title"), "Edited original title")
		if err != nil || !changed {
			t.Fatalf("held anchor was invalidated: %v", err)
		}
		save := func(path string) error { _, e := w.session.SaveAs(path); return e }
		contract20CreationSaveFaults(t, save)
		dest := filepath.Join(t.TempDir(), "result.pptx")
		if err := save(dest); err != nil {
			t.Fatal(err)
		}
		_, after, graph := contract20CreationArchive(t, dest)
		if !bytes.Equal(w.source, w.original) {
			t.Fatal("source changed")
		}
		base, err := blankMemberPayloads(w.baseFixture)
		if err != nil {
			t.Fatal(err)
		}
		if !reflect.DeepEqual(base, w.before) {
			t.Fatal("sealed fixture or derived source changed")
		}
		if len(after) != len(w.before)+2 {
			t.Fatalf("member delta: %d -> %d", len(w.before), len(after))
		}
		allowed := map[string]bool{"ppt/slides/slide1.xml": true, "ppt/presentation.xml": true, "ppt/_rels/presentation.xml.rels": true, "[Content_Types].xml": true}
		for name, b := range w.before {
			if got, ok := after[name]; !ok || (!allowed[name] && !bytes.Equal(b, got)) {
				t.Fatalf("unrelated member changed/removed: %s", name)
			}
		}
		for name := range after {
			if _, ok := w.before[name]; !ok && name != "ppt/slides/slide2.xml" && name != "ppt/slides/_rels/slide2.xml.rels" {
				t.Fatalf("unrequested added member %s", name)
			}
		}
		if !bytes.Equal(w.before["ppt/slides/_rels/slide1.xml.rels"], after["ppt/slides/_rels/slide1.xml.rels"]) {
			t.Fatal("old layout edge changed")
		}
		contract20AppendGraph(t, w.graph, graph, "ppt/slides/slide1.xml", "ppt/slides/slide2.xml")
		if err := creationAppendInsertions(w.before, after, w.graph, graph, "ppt/slides/slide1.xml", "ppt/slides/slide2.xml"); err != nil {
			t.Fatal(err)
		}
		for name := range after {
			if _, ok := w.before[name]; !ok && name == "ppt/slides/_rels/slide2.xml.rels" {
				if !bytes.Contains(after[name], []byte("slideLayout")) {
					t.Fatal("new relationship registry lacks layout")
				}
			}
		}
		old := w.before["ppt/slides/slide1.xml"]
		want := bytes.Replace(old, []byte("Original title"), []byte("Edited original title"), 1)
		if bytes.Equal(old, want) || !bytes.Equal(want, after["ppt/slides/slide1.xml"]) {
			t.Fatal("old slide changed beyond selected title span")
		}
		if !bytes.Contains(after["ppt/slides/slide2.xml"], []byte(">Appended title</a:t>")) || !bytes.Contains(after["ppt/slides/slide2.xml"], []byte(">Appended subtitle</a:t>")) {
			t.Fatal("new placeholder operands")
		}
		if reflect.DeepEqual(graph, w.graph) {
			t.Fatal("append graph not extended")
		}
		deck, err := presentation.Open(dest)
		if err != nil {
			t.Fatal(err)
		}
		defer deck.Close()
		if deck.SlideCount() != 2 {
			t.Fatal("reopened two slides required")
		}
		for i, pair := range [][2]string{{"Edited original title", "Original subtitle"}, {"Appended title", "Appended subtitle"}} {
			slide, e := deck.Slide(i + 1)
			if e != nil {
				t.Fatal(e)
			}
			texts := []string{}
			for _, shape := range slide.Shapes() {
				texts = append(texts, shape.Text())
			}
			if !reflect.DeepEqual(texts, []string{pair[0], pair[1]}) {
				t.Fatalf("reopened slide %d: %q", i, texts)
			}
		}
	})
	t.Run("refusals", func(t *testing.T) {
		contract20CreationSelected(t, "@id-pptx-create-refusals")
		creator, err := presentation.NewContractPresentation()
		if err != nil {
			t.Fatal(err)
		}
		defer creator.Close()
		contract20CreationRefusal(t, creator.AddTitleSlide("bad\x00", "Bad subtitle"), "PPTX_ARGUMENT_INVALID")
		if creator.SlideCount() != 0 {
			t.Fatal("invalid title mutated creator")
		}
		absent := filepath.Join(t.TempDir(), "refused.pptx")
		contract20CreationRefusal(t, creator.SaveAs(absent), "PPTX_CREATION_UNSUPPORTED")
		if _, err := os.Stat(absent); !os.IsNotExist(err) {
			t.Fatalf("creator refusal published destination: %v", err)
		}
		w := &contract20TitleWorld{temp: t.TempDir()}
		if err := w.input("unsafe-title-layout"); err != nil {
			t.Fatal(err)
		}
		anchor, err := w.session.FindContractTitle("ppt/slides/slide1.xml")
		if err != nil {
			t.Fatal(err)
		}
		contract20CreationRefusal(t, w.session.AppendContractTitleSlide("Unsafe next title", "Unsafe next subtitle"), "PPTX_LAYOUT_UNSAFE")
		if !bytes.Equal(w.source, w.original) {
			t.Fatal("caller input changed")
		}
		dest := filepath.Join(t.TempDir(), "unchanged.pptx")
		if _, err := w.session.SaveAs(dest); err != nil {
			t.Fatal(err)
		}
		_, members, graph := contract20CreationArchive(t, dest)
		if !reflect.DeepEqual(members, w.before) || !reflect.DeepEqual(graph, w.graph) || anchor.Text() != "Original title" {
			t.Fatal("refusal mutated members, graph or held read")
		}
		before := bytes.Clone(members["ppt/slides/slide1.xml"])
		changed, err := w.session.ReplaceContractTitleSpan(anchor, 0, len("Original title"), "Retained original title")
		if err != nil || !changed {
			t.Fatalf("held anchor unusable after unsafe refusal: %v", err)
		}
		recovered := filepath.Join(t.TempDir(), "recovered.pptx")
		if _, err := w.session.SaveAs(recovered); err != nil {
			t.Fatal(err)
		}
		_, reopened, _ := contract20CreationArchive(t, recovered)
		if bytes.Equal(before, reopened["ppt/slides/slide1.xml"]) || len(reopened) != len(w.before) {
			t.Fatal("recovered edit or member inventory")
		}
		if err := creator.AddTitleSlide("Recovered title", "Recovered subtitle"); err != nil {
			t.Fatalf("creator unusable after refusal: %v", err)
		}
		if creator.SlideCount() != 1 {
			t.Fatal("creator recovered count")
		}
	})
}
