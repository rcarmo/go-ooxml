package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"path/filepath"
	"testing"
)

func graphEditPackage(t *testing.T) *Preserved {
	t.Helper()
	q := New()
	_, _ = q.AddPart("word/document.xml", ContentTypeXML, []byte(`<r/>`))
	_, _ = q.AddPart("word/media/shared.bin", "application/octet-stream", []byte("shared"))
	_, _ = q.AddPart("opaque.bin", "application/octet-stream", []byte("opaque"))
	q.AddRelationship("", "word/document.xml", RelTypeOfficeDocument)
	q.AddRelationship("word/document.xml", "media/shared.bin", RelTypeImage)
	q.AddRelationship("word/document.xml", "media/shared.bin", RelTypeImage)
	var b bytes.Buffer
	if err := q.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	p, err := OpenPreserved(b.Bytes(), Limits{})
	if err != nil {
		t.Fatal(err)
	}
	return p
}
func graphEditRequest() GraphMutation {
	return GraphMutation{Additions: []PartAddition{{"word/media/new.bin", "application/octet-stream", []byte("new")}}, Retargets: []RelationshipRetarget{{"word/document.xml", "rId1", "word/media/new.bin"}}}
}
func serializeGraphEdit(t *testing.T, p *Preserved) []byte {
	t.Helper()
	var b bytes.Buffer
	if err := p.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	return b.Bytes()
}
func TestGraphPlanContracts(t *testing.T) {
	t.Run("private plan and payload snapshots", func(t *testing.T) {
		p := graphEditPackage(t)
		c := graphEditRequest()
		plan, err := p.PlanGraphMutation(c)
		if err != nil {
			t.Fatal(err)
		}
		c.Additions[0].Data[0] = 'X'
		c.Retargets[0].TargetPart = "absent"
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
		data, hash, err := p.Part("word/media/new.bin")
		if err != nil || string(data) != "new" {
			t.Fatal(string(data), err)
		}
		data[0] = 'Y'
		if err = p.Replace([]Replacement{{"word/media/new.bin", hash, []byte("updated")}}); err != nil {
			t.Fatal(err)
		}
		if _, err = p.Graph(); err != nil {
			t.Fatal(err)
		}
		out := serializeGraphEdit(t, p)
		if !bytes.Equal(out, serializeGraphEdit(t, p)) {
			t.Fatal("serialization nondeterministic")
		}
		r, err := p.SaveAs(filepath.Join(t.TempDir(), "out.zip"))
		if err != nil || r.Schema != 2 {
			t.Fatal(r, err)
		}
		for _, change := range r.Changes {
			if change.Part == "word/media/new.bin" && (change.Operation != "add" || change.BeforeSHA256 != "" || change.AfterSHA256 != fingerprint([]byte("updated"))) {
				t.Fatal(change)
			}
		}
	})
	t.Run("foreign consumed and ABA stale plans", func(t *testing.T) {
		p := graphEditPackage(t)
		plan, err := p.PlanGraphMutation(graphEditRequest())
		if err != nil {
			t.Fatal(err)
		}
		other := graphEditPackage(t)
		if other.ApplyGraphPlan(plan) == nil {
			t.Fatal("foreign plan accepted")
		}
		old, hash, _ := p.Part("opaque.bin")
		if err = p.Replace([]Replacement{{"opaque.bin", hash, []byte("change")}}); err != nil {
			t.Fatal(err)
		}
		_, hash, _ = p.Part("opaque.bin")
		if err = p.Replace([]Replacement{{"opaque.bin", hash, old}}); err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err == nil {
			t.Fatal("ABA stale accepted")
		}
		plan, err = p.PlanGraphMutation(graphEditRequest())
		if err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err == nil {
			t.Fatal("consumed plan accepted")
		}
	})
	t.Run("no-op keeps exact source and plan", func(t *testing.T) {
		p := graphEditPackage(t)
		original := serializeGraphEdit(t, p)
		plan, err := p.PlanGraphMutation(GraphMutation{Retargets: []RelationshipRetarget{{"word/document.xml", "rId1", "word/media/shared.bin"}}})
		if err != nil {
			t.Fatal(err)
		}
		for i := 0; i < 2; i++ {
			if err = p.ApplyGraphPlan(plan); err != nil {
				t.Fatal(err)
			}
		}
		if !bytes.Equal(original, serializeGraphEdit(t, p)) {
			t.Fatal("no-op reserialized")
		}
	})
	t.Run("raw unchanged members", func(t *testing.T) {
		p := graphEditPackage(t)
		original := serializeGraphEdit(t, p)
		plan, err := p.PlanGraphMutation(graphEditRequest())
		if err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
		out := serializeGraphEdit(t, p)
		a, _ := zip.NewReader(bytes.NewReader(original), int64(len(original)))
		b, _ := zip.NewReader(bytes.NewReader(out), int64(len(out)))
		files := map[string]*zip.File{}
		for _, f := range b.File {
			files[f.Name] = f
		}
		for _, f := range a.File {
			if f.Name == ContentTypesPath || f.Name == RelationshipsPathForPart("word/document.xml") {
				continue
			}
			g := files[f.Name]
			if g == nil {
				t.Fatal("original removed")
			}
			ar, _ := f.OpenRaw()
			br, _ := g.OpenRaw()
			ab, _ := io.ReadAll(ar)
			bb, _ := io.ReadAll(br)
			if !bytes.Equal(ab, bb) || !bytes.Equal(f.Extra, g.Extra) || f.Comment != g.Comment || f.ExternalAttrs != g.ExternalAttrs {
				t.Fatal("untouched raw member changed", f.Name)
			}
		}
	})
	t.Run("all additions preflighted and empty added data remains present", func(t *testing.T) {
		p := graphEditPackage(t)
		original := serializeGraphEdit(t, p)
		c := graphEditRequest()
		c.Additions = append(c.Additions, PartAddition{"bad.bin", "text/xml", []byte(`<bad>`)})
		if _, err := p.PlanGraphMutation(c); err == nil {
			t.Fatal("late malformed addition accepted")
		}
		if !bytes.Equal(original, serializeGraphEdit(t, p)) {
			t.Fatal("partial plan commit")
		}
		c = GraphMutation{Additions: []PartAddition{{"empty.bin", "application/octet-stream", nil}}}
		plan, err := p.PlanGraphMutation(c)
		if err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
		if _, _, err = p.Part("empty.bin"); err != nil {
			t.Fatal(err)
		}
	})
	t.Run("signature and URI-policy refusals", func(t *testing.T) {
		for _, name := range []string{"../outside", "/absolute.bin", "a%20b.bin", "a./b.bin", "_XMLSIGNATURES/sig.bin", "a/_rels/file.xml.rels"} {
			p := graphEditPackage(t)
			c := graphEditRequest()
			c.Additions[0].Name = name
			_, err := p.PlanGraphMutation(c)
			var r *Refusal
			if !errors.As(err, &r) {
				t.Fatal(name, err)
			}
		}
		p := graphEditPackage(t)
		p.signed = true
		if _, err := p.PlanGraphMutation(graphEditRequest()); err == nil {
			t.Fatal("signed edit accepted")
		}
	})
}
