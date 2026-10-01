package document

import (
	"bytes"
	"errors"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"os"
	"path/filepath"
	"testing"
)

func TestRetainedTableWordNative(t *testing.T) {
	const fixture = "fixtures/docx/table-properties/table-properties-ea4dfb625197.docx"
	const part = "word/document.xml"
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
		{"shading", map[string]any{"shading": "A1B2C3"}, `w:fill="A1B2C3"`},
		{"remove shading", map[string]any{"shading": nil}, `<w:tcBorders>`},
		{"alignment", map[string]any{"verticalAlign": "center"}, `<w:vAlign w:val="center"/>`},
		{"direction", map[string]any{"textDirection": "tbRl"}, `<w:textDirection w:val="tbRl"/>`},
		{"margins", map[string]any{"margins": map[string]any{"top": 120, "left": 240, "bottom": 120, "right": 240}}, `<w:top w:w="120" w:type="dxa"/>`},
		{"top border", map[string]any{"border": map[string]any{"side": "top", "style": "single", "size": 8, "color": "334455"}}, `<w:top w:val="single" w:sz="8" w:color="334455"/>`},
		{"left border", map[string]any{"border": map[string]any{"side": "left", "style": "single", "size": 8, "color": "334455"}}, `<w:left w:val="single" w:sz="8" w:color="334455"/>`},
		{"bottom border", map[string]any{"border": map[string]any{"side": "bottom", "style": "single", "size": 8, "color": "334455"}}, `<w:bottom w:val="single" w:sz="8" w:color="334455"/>`},
		{"right border", map[string]any{"border": map[string]any{"side": "right", "style": "single", "size": 8, "color": "334455"}}, `<w:right w:val="single" w:sz="8" w:color="334455"/>`},
		{"remove border", map[string]any{"border": map[string]any{"side": "top", "remove": true}}, `<w:left w:val="single"`},
		{"no wrap", map[string]any{"noWrap": true}, `<w:noWrap w:val="1"/>`},
		{"fit text", map[string]any{"tcFitText": true}, `<w:tcFitText w:val="1"/>`},
		{"preferred width", map[string]any{"width": 2400}, `<w:tcW w:w="2400" w:type="dxa"/>`},
		{"header", map[string]any{"tblHeader": true}, `<w:tblHeader w:val="1"/>`},
		{"cant split", map[string]any{"cantSplit": true}, `<w:cantSplit w:val="1"/>`},
		{"height", map[string]any{"height": 480, "heightRule": "exact"}, `<w:trHeight w:val="480" w:hRule="exact"/>`},
		{"table alignment", map[string]any{"alignment": "center"}, `<w:jc w:val="center"/>`},
		{"indent", map[string]any{"indent": 720}, `<w:tblInd w:w="720" w:type="dxa"/>`},
		{"hide mark", map[string]any{"hideMark": true}, `<w:hideMark w:val="1"/>`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			caller := bytes.Clone(original)
			s, er := OpenEditing(original, packaging.Limits{})
			if er != nil {
				t.Fatal(er)
			}
			target, er := s.FindRetainedTable(0, 0, 0)
			if er != nil {
				t.Fatal(er)
			}
			if er = s.SetRetainedTable(target, c.patch); er != nil {
				t.Fatal(er)
			}
			dest := filepath.Join(t.TempDir(), "saved.docx")
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
		target, e := s.FindRetainedTable(0, 0, 0)
		if e != nil {
			t.Fatal(e)
		}
		e = s.SetRetainedTable(target, map[string]any{"shading": "A1B2C3", "border": map[string]any{"side": "top", "style": "single", "size": 97, "color": "334455"}})
		var r *packaging.Refusal
		if !errors.As(e, &r) || r.Kind != "invalid-table-properties" {
			t.Fatalf("refusal: %v", e)
		}
		dest := filepath.Join(t.TempDir(), "unchanged.docx")
		if _, e = s.SaveAs(dest); e != nil {
			t.Fatal(e)
		}
		got, e := os.ReadFile(dest)
		if e != nil || !bytes.Equal(got, original) {
			t.Fatalf("partial refusal %v", e)
		}
		if e = s.SetRetainedTable(target, map[string]any{"shading": "4472C4"}); e != nil {
			t.Fatalf("held target stale: %v", e)
		}
	})
}
