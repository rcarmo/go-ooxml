package acceptance

import (
	"bytes"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Relocated slide paths are a native oracle/control, not another canonical
// case. The producer may allocate any unique names; tests resolve list order
// and relationship targets rather than silently blessing numbered members.
func TestContract20CreationRelocatedPartControl(t *testing.T) {
	contract20CreationSelected(t, "@id-pptx-create-anchor-preservation")
	w := &contract20TitleWorld{temp: t.TempDir()}
	if e := w.input("title-subtitle-source"); e != nil {
		t.Fatal(e)
	}
	before := map[string][]byte{}
	for name, data := range w.before {
		before[name] = bytes.Clone(data)
	}
	old := "ppt/slides/slide1.xml"
	renamed := "ppt/slides/sourceTitle.xml"
	oldRels := packaging.RelationshipsPathForPart(old)
	newRels := packaging.RelationshipsPathForPart(renamed)
	if _, ok := before[renamed]; ok {
		t.Fatal("relocated name already exists")
	}
	before[renamed] = before[old]
	delete(before, old)
	before[newRels] = before[oldRels]
	delete(before, oldRels)
	rels := before["ppt/_rels/presentation.xml.rels"]
	if bytes.Count(rels, []byte(`Target="slides/slide1.xml"`)) != 1 {
		t.Fatal("source slide edge not unique")
	}
	before["ppt/_rels/presentation.xml.rels"] = bytes.Replace(rels, []byte(`Target="slides/slide1.xml"`), []byte(`Target="slides/sourceTitle.xml"`), 1)
	registry := before["[Content_Types].xml"]
	if bytes.Count(registry, []byte(`/ppt/slides/slide1.xml"`)) != 1 {
		t.Fatal("source override not unique")
	}
	before["[Content_Types].xml"] = bytes.Replace(registry, []byte(`/ppt/slides/slide1.xml"`), []byte(`/ppt/slides/sourceTitle.xml"`), 1)
	source, e := contract20PackMembers(before)
	if e != nil {
		t.Fatal(e)
	}
	graph, e := contract20TableGraph(source)
	if e != nil {
		t.Fatal(e)
	}
	order, e := creationSlideOrder(before, graph)
	if e != nil || !reflect.DeepEqual(order, []string{renamed}) {
		t.Fatalf("relocated source order %q: %v", order, e)
	}
	session, e := presentation.OpenEditing(source, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	anchor, e := session.FindContractTitle(order[0])
	if e != nil || anchor.Text() != "Original title" {
		t.Fatalf("relocated title anchor: %v", e)
	}
	if e = session.AppendContractTitleSlide("Appended title", "Appended subtitle"); e != nil {
		t.Fatal(e)
	}
	changed, e := session.ReplaceContractTitleSpan(anchor, 0, len("Original title"), "Edited original title")
	if e != nil || !changed {
		t.Fatalf("relocated held anchor: %v", e)
	}
	dest := filepath.Join(t.TempDir(), "relocated.pptx")
	if _, e = session.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	saved, e := os.ReadFile(dest)
	if e != nil {
		t.Fatal(e)
	}
	after, e := blankMemberPayloads(saved)
	if e != nil {
		t.Fatal(e)
	}
	next, e := contract20TableGraph(saved)
	if e != nil {
		t.Fatal(e)
	}
	out, e := creationSlideOrder(after, next)
	if e != nil || len(out) != 2 || out[0] != renamed || out[1] == renamed {
		t.Fatalf("relocated output order %q: %v", out, e)
	}
	if e = creationAppendInsertions(before, after, graph, next, renamed, out[1]); e != nil {
		t.Fatal(e)
	}
	if e = creationGraphDelta(graph, next, renamed, out[1]); e != nil {
		t.Fatal(e)
	}
	if !bytes.Equal(after[newRels], before[newRels]) {
		t.Fatal("relocated old relationship changed")
	}
	original := before[renamed]
	expected := bytes.Replace(original, []byte("Original title"), []byte("Edited original title"), 1)
	if bytes.Equal(original, expected) || !bytes.Equal(after[renamed], expected) {
		t.Fatal("relocated slide changed outside selected title text")
	}
	for i, pair := range [][2]string{{"Edited original title", "Original subtitle"}, {"Appended title", "Appended subtitle"}} {
		if e = creationPlaceholderOperands(after[out[i]], pair[0], pair[1]); e != nil {
			t.Fatal(e)
		}
	}
	if e = creationReopened(dest, [][2]string{{"Edited original title", "Original subtitle"}, {"Appended title", "Appended subtitle"}}); e != nil {
		t.Fatal(e)
	}
	// The original caller archive remains under its original member identity.
	if !strings.Contains(string(w.before["ppt/_rels/presentation.xml.rels"]), `slides/slide1.xml`) {
		t.Fatal("fixture source changed")
	}
}
