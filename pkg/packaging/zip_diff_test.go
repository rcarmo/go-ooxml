package packaging

import (
	"archive/zip"
	"bytes"
	"reflect"
	"testing"
)

func rawDiffArchive(t *testing.T, entries []ZIP32Entry) []byte {
	t.Helper()
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, e := range entries {
		w, err := zw.CreateHeader(&zip.FileHeader{Name: e.Name, Method: zip.Store})
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(e.Data); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}
func TestRawZIP32SemanticDiffFamily(t *testing.T) {
	old := rawDiffArchive(t, []ZIP32Entry{{Name: "a.xml", Data: []byte(`<a xmlns="urn:x"/>`)}, {Name: "b.bin", Data: []byte("old")}})
	newer := rawDiffArchive(t, []ZIP32Entry{{Name: "a.xml", Data: []byte(`<p:a xmlns:p="urn:x"/>`)}, {Name: "b.bin", Data: []byte("new")}, {Name: "c.bin", Data: []byte("added")}})
	first, second := bytes.Clone(old), bytes.Clone(newer)
	diff, err := DiffZIP32(old, newer, ZIP32Limits{})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(diff.EquivalentXML, []string{"a.xml"}) || !reflect.DeepEqual(diff.Changed, []string{"b.bin"}) || !reflect.DeepEqual(diff.Added, []string{"c.bin"}) || len(diff.Removed) != 0 {
		t.Fatalf("raw diff: %+v", diff)
	}
	if !bytes.Equal(old, first) || !bytes.Equal(newer, second) {
		t.Fatal("input changed")
	}
	diff.Added[0] = "mutated"
	again, err := DiffZIP32(old, newer, ZIP32Limits{})
	if err != nil || again.Added[0] != "c.bin" {
		t.Fatal("diff aliases state", err)
	}
	unsafe := append(bytes.Clone(newer), '\n')
	if result, err := DiffZIP32(old, unsafe, ZIP32Limits{}); err == nil || len(result.Added)+len(result.Removed)+len(result.Changed)+len(result.EquivalentXML) != 0 {
		t.Fatal("unsafe archive compared")
	}
	// Identical unsafe XML is not an equivalence shortcut; the report must
	// classify a non-identical member as changed, not equivalent XML.
	left := rawDiffArchive(t, []ZIP32Entry{{Name: "a.xml", Data: []byte(`<!DOCTYPE a><a/>`)}})
	right := rawDiffArchive(t, []ZIP32Entry{{Name: "a.xml", Data: []byte(`<!DOCTYPE a><a x="1"/>`)}})
	report, err := DiffZIP32(left, right, ZIP32Limits{})
	if err != nil || !reflect.DeepEqual(report.Changed, []string{"a.xml"}) || len(report.EquivalentXML) != 0 {
		t.Fatalf("unsafe XML falsely equivalent: %+v %v", report, err)
	}
	if !bytes.Equal(old, first) || !bytes.Equal(newer, second) {
		t.Fatal("diff fault changed input")
	}
}
