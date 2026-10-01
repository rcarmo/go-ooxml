package presentation

import (
	"bytes"
	"errors"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"os"
	"path/filepath"
	"testing"
)

func TestRetainedTablePPTXNative(t *testing.T) {
	const fixture = "fixtures/pptx/table-properties/table-properties-ce68a5cbf3d3.pptx"
	const part = "ppt/slides/slide1.xml"
	original, e := os.ReadFile(testutil.ReferencePath(filepath.FromSlash(fixture)))
	if e != nil {
		if os.IsNotExist(e) && os.Getenv("OOXML_FIXTURES_ROOT") == "" && os.Getenv("OOXML_REFERENCE_PIN") == "" {
			t.Skip("candidate absent from default pin")
		}
		t.Fatal(e)
	}
	cases := []struct {
		name  string
		patch map[string]any
		want  string
	}{
		{"solid fill", map[string]any{"fill": "A1B2C3"}, `<a:srgbClr val="A1B2C3"/>`},
		{"no fill", map[string]any{"fill": "none"}, `<a:noFill/>`},
		{"remove fill", map[string]any{"fill": nil}, `<a:tcPr marL=`},
		{"border left", map[string]any{"border": map[string]any{"side": "left", "color": "334455", "width": 25400}}, `<a:lnL w="25400">`},
		{"border right", map[string]any{"border": map[string]any{"side": "right", "color": "334455", "width": 25400}}, `<a:lnR w="25400">`},
		{"border top", map[string]any{"border": map[string]any{"side": "top", "color": "334455", "width": 25400}}, `<a:lnT w="25400">`},
		{"border bottom", map[string]any{"border": map[string]any{"side": "bottom", "color": "334455", "width": 25400}}, `<a:lnB w="25400">`},
		{"border nofill", map[string]any{"border": map[string]any{"side": "top", "color": "none"}}, `<a:lnT w="12700"><a:noFill/>`},
		{"border remove", map[string]any{"border": map[string]any{"side": "left", "remove": true}}, `<a:lnR w="12700">`},
		{"margins", map[string]any{"margins": map[string]any{"left": 91440, "right": 91440, "top": 45720, "bottom": 45720}}, `marL="91440"`},
		{"anchor", map[string]any{"anchor": "ctr"}, `anchor="ctr"`},
		{"direction", map[string]any{"vert": "vert"}, `vert="vert"`},
		{"anchor center", map[string]any{"anchorCtr": true}, `anchorCtr="1"`},
		{"column width", map[string]any{"column": 0, "width": 3657600}, `cx="9144000"`},
		{"row height", map[string]any{"row": 0, "height": 914400}, `cy="3962400"`},
		{"first row", map[string]any{"firstRow": false}, `firstRow="0"`},
		{"band row", map[string]any{"bandRow": false}, `bandRow="0"`},
		{"frame position", map[string]any{"x": 914400, "y": 1828800}, `<a:off x="914400" y="1828800"/>`},
		{"unbold", map[string]any{"bold": false}, `<a:rPr b="0"/>`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			caller := bytes.Clone(original)
			s, er := OpenEditing(original, packaging.Limits{})
			if er != nil {
				t.Fatal(er)
			}
			target, er := s.FindRetainedTable(part, 3, 0, 0)
			if er != nil {
				t.Fatal(er)
			}
			if er = s.SetRetainedTable(target, c.patch); er != nil {
				t.Fatal(er)
			}
			dest := filepath.Join(t.TempDir(), "saved.pptx")
			if _, er = s.SaveAs(dest); er != nil {
				t.Fatal(er)
			}
			output, er := os.ReadFile(dest)
			if er != nil {
				t.Fatal(er)
			}
			q, er := packaging.OpenPreserved(output, packaging.Limits{})
			if er != nil {
				t.Fatal(er)
			}
			xml, _, er := q.Part(part)
			if er != nil {
				t.Fatal(er)
			}
			if !bytes.Contains(xml, []byte(c.want)) {
				t.Fatalf("missing %q", c.want)
			}
			if bytes.Equal(output, original) || !bytes.Equal(original, caller) {
				t.Fatal("edit lost or caller mutated")
			}
		})
	}
	t.Run("late invalid refusal", func(t *testing.T) {
		s, e := OpenEditing(original, packaging.Limits{})
		if e != nil {
			t.Fatal(e)
		}
		target, e := s.FindRetainedTable(part, 3, 0, 0)
		if e != nil {
			t.Fatal(e)
		}
		e = s.SetRetainedTable(target, map[string]any{"fill": "A1B2C3", "border": map[string]any{"side": "top", "color": "334455", "width": -1}})
		var r *packaging.Refusal
		if !errors.As(e, &r) || r.Kind != "invalid-table-properties" {
			t.Fatalf("refusal: %v", e)
		}
		dest := filepath.Join(t.TempDir(), "unchanged.pptx")
		if _, e = s.SaveAs(dest); e != nil {
			t.Fatal(e)
		}
		got, e := os.ReadFile(dest)
		if e != nil || !bytes.Equal(got, original) {
			t.Fatalf("partial refusal %v", e)
		}
		if e = s.SetRetainedTable(target, map[string]any{"fill": "4472C4"}); e != nil {
			t.Fatalf("held target stale: %v", e)
		}
	})
}
