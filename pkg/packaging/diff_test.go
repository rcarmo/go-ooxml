package packaging

import (
	"bytes"
	"reflect"
	"testing"
)

func TestDiffPreservedEffectiveTypesAndPayloads(t *testing.T) {
	first := graphEditPackage(t)
	second := graphEditPackage(t)
	firstArchive := serializeGraphEdit(t, first)
	plan, err := second.PlanGraphMutation(GraphMutation{ContentTypes: []ContentTypeChange{{Part: "word/document.xml", ContentType: "application/vnd.test.document+xml"}}})
	if err != nil {
		t.Fatal(err)
	}
	if err = second.ApplyGraphPlan(plan); err != nil {
		t.Fatal(err)
	}
	report, err := DiffPreserved(first, second)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(report.Changed, []string{ContentTypesPath, "word/document.xml"}) || len(report.Added) != 0 || len(report.Removed) != 0 || len(report.EquivalentXML) != 0 {
		t.Fatalf("type-only diff: %+v", report)
	}
	oldPayload, _, _ := first.Part("word/document.xml")
	newPayload, _, _ := second.Part("word/document.xml")
	if !bytes.Equal(oldPayload, newPayload) || !bytes.Equal(firstArchive, serializeGraphEdit(t, first)) {
		t.Fatal("type-only payload or original snapshot changed")
	}

	a := New()
	_, _ = a.AddPart("a.xml", ContentTypeXML, []byte(`<p:r xmlns:p="urn:test" a="1"/>`))
	_, _ = a.AddPart("b.bin", "application/octet-stream", []byte("old"))
	b := New()
	_, _ = b.AddPart("a.xml", ContentTypeXML, []byte(`<q:r a="1" xmlns:q="urn:test"/>`))
	_, _ = b.AddPart("b.bin", "application/octet-stream", []byte("new"))
	_, _ = b.AddPart("c.bin", "application/octet-stream", []byte("added"))
	open := func(q *Package) *Preserved {
		var out bytes.Buffer
		if err := q.WriteTo(&out); err != nil {
			t.Fatal(err)
		}
		p, err := OpenPreserved(out.Bytes(), Limits{})
		if err != nil {
			t.Fatal(err)
		}
		return p
	}
	pa, pb := open(a), open(b)
	compared, err := DiffPreserved(pa, pb)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(compared.Added, []string{"c.bin"}) || !reflect.DeepEqual(compared.Changed, []string{ContentTypesPath, "b.bin"}) || !reflect.DeepEqual(compared.EquivalentXML, []string{"a.xml"}) || len(compared.Removed) != 0 {
		t.Fatalf("semantic raw diff: %+v", compared)
	}
	compared.Added[0] = "mutated"
	fresh, err := DiffPreserved(pa, pb)
	if err != nil || fresh.Added[0] != "c.bin" {
		t.Fatal("returned diff aliases state", err)
	}
}
