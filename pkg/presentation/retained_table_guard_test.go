package presentation

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func tableGuardFixture(t *testing.T) []byte {
	t.Helper()
	b, e := os.ReadFile(testutil.ReferencePath("fixtures/pptx/table-properties/table-properties-ce68a5cbf3d3.pptx"))
	if os.IsNotExist(e) && os.Getenv("OOXML_FIXTURES_ROOT") == "" && os.Getenv("OOXML_REFERENCE_PIN") == "" {
		t.Skip("candidate absent from default pin")
	}
	if e != nil {
		t.Fatal(e)
	}
	return b
}
func tableGuardChanged(t *testing.T, input []byte, old, new string) []byte {
	t.Helper()
	s, e := OpenEditing(input, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	body, hash, e := s.pkg.Part("ppt/slides/slide1.xml")
	if e != nil {
		t.Fatal(e)
	}
	if bytes.Count(body, []byte(old)) != 1 {
		t.Fatalf("fixture change %q not unique: %d", old, bytes.Count(body, []byte(old)))
	}
	changed := bytes.Replace(body, []byte(old), []byte(new), 1)
	if e = s.pkg.Replace([]packaging.Replacement{{Part: "ppt/slides/slide1.xml", ExpectedSHA256: hash, Data: changed}}); e != nil {
		t.Fatal(e)
	}
	dest := filepath.Join(t.TempDir(), "mutated.pptx")
	if _, e = s.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	result, e := os.ReadFile(dest)
	if e != nil {
		t.Fatal(e)
	}
	return result
}
func TestRetainedTablePPTXGuardBatch(t *testing.T) {
	input := tableGuardFixture(t)
	cases := []struct {
		name, old, new string
		patch          map[string]any
	}{
		{"hyperlink", `<a:t>Chapter</a:t>`, `<a:hlinkClick r:id="rId9"/><a:t>Chapter</a:t>`, map[string]any{"fill": "A1B2C3"}},
		{"field", `<a:t>Chapter</a:t>`, `<a:fld id="f1" type="slidenum"><a:t>Chapter</a:t></a:fld>`, map[string]any{"fill": "A1B2C3"}},
		{"frame rotation", `<p:xfrm><a:off x="457200"`, `<p:xfrm rot="60000"><a:off x="457200"`, map[string]any{"fill": "A1B2C3"}},
		{"graphic frame locked", `<a:graphicFrameLocks noGrp="1"/>`, `<a:graphicFrameLocks noGrp="1" noMove="1"/>`, map[string]any{"fill": "A1B2C3"}},
		{"unexpected frame child", `</p:nvGraphicFramePr><p:xfrm>`, `</p:nvGraphicFramePr><p:extLst/><p:xfrm>`, map[string]any{"fill": "A1B2C3"}},
		{"reordered table owners", `</a:tblPr><a:tblGrid>`, `</a:tblPr><a:tr h="609600"/><a:tblGrid>`, map[string]any{"fill": "A1B2C3"}},
		{"unknown grid child", `<a:tblGrid><a:gridCol w="2743200"/>`, `<a:tblGrid><a:bad/><a:gridCol w="2743200"/>`, map[string]any{"fill": "A1B2C3"}},
		{"unknown row child", `</a:tblGrid><a:tr h="609600"><a:tc>`, `</a:tblGrid><a:tr h="609600"><a:bad/><a:tc>`, map[string]any{"fill": "A1B2C3"}},
		{"unknown cell child", `</a:tblGrid><a:tr h="609600"><a:tc><a:txBody>`, `</a:tblGrid><a:tr h="609600"><a:tc><a:bad/><a:txBody>`, map[string]any{"fill": "A1B2C3"}},
		{"short selected row", `</a:tblGrid><a:tr h="609600"><a:tc>`, `</a:tblGrid><a:tr h="609600"><a:tc gridSpan="2">`, map[string]any{"fill": "A1B2C3"}},
		{"merged neighbor", `</a:tblGrid><a:tr h="609600"><a:tc><a:txBody><a:bodyPr/>`, `</a:tblGrid><a:tr h="609600"><a:tc gridSpan="2"><a:txBody><a:bodyPr/>`, map[string]any{"fill": "A1B2C3"}},
		{"invalid original flag", `<a:tblPr firstRow="1" bandRow="1">`, `<a:tblPr firstRow="maybe" bandRow="1">`, map[string]any{"fill": "A1B2C3"}},
		{"invalid original margin", `<a:tcPr marL="0"`, `<a:tcPr marL="bad"`, map[string]any{"fill": "A1B2C3"}},
		{"wrong extent sum", `<a:ext cx="8229600" cy="3657600"/>`, `<a:ext cx="8229601" cy="3657600"/>`, map[string]any{"fill": "A1B2C3"}},
		{"near-bound width growth", `<a:off x="457200" y="1371600"/>`, `<a:off x="91770400" y="1371600"/>`, map[string]any{"column": 0, "width": 3657600}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			mutated := tableGuardChanged(t, input, c.old, c.new)
			s, e := OpenEditing(mutated, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target, e := s.FindRetainedTable("ppt/slides/slide1.xml", 3, 0, 0)
			if e == nil {
				e = s.SetRetainedTable(target, c.patch)
			}
			var refusal *packaging.Refusal
			if !errors.As(e, &refusal) {
				t.Fatalf("not refused: %v", e)
			}
			dest := filepath.Join(t.TempDir(), "refused.pptx")
			if _, e = s.SaveAs(dest); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(dest)
			if e != nil || !bytes.Equal(saved, mutated) {
				t.Fatalf("refusal changed archive %v", e)
			}
		})
	}
}
