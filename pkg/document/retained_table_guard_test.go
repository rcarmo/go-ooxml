package document

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func wordTableGuardFixture(t *testing.T) []byte {
	t.Helper()
	b, e := os.ReadFile(testutil.ReferencePath("fixtures/docx/table-properties/table-properties-ea4dfb625197.docx"))
	if os.IsNotExist(e) && os.Getenv("OOXML_FIXTURES_ROOT") == "" && os.Getenv("OOXML_REFERENCE_PIN") == "" {
		t.Skip("candidate absent from default pin")
	}
	if e != nil {
		t.Fatal(e)
	}
	return b
}
func wordTableGuardChanged(t *testing.T, input []byte, old, new string) []byte {
	t.Helper()
	s, e := OpenEditing(input, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	body, hash, e := s.pkg.Part("word/document.xml")
	if e != nil {
		t.Fatal(e)
	}
	if bytes.Count(body, []byte(old)) != 1 {
		t.Fatalf("fixture change %q not unique: %d", old, bytes.Count(body, []byte(old)))
	}
	changed := bytes.Replace(body, []byte(old), []byte(new), 1)
	if e = s.pkg.Replace([]packaging.Replacement{{Part: "word/document.xml", ExpectedSHA256: hash, Data: changed}}); e != nil {
		t.Fatal(e)
	}
	dest := filepath.Join(t.TempDir(), "mutated.docx")
	if _, e = s.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	result, e := os.ReadFile(dest)
	if e != nil {
		t.Fatal(e)
	}
	return result
}
func TestRetainedTableWordGuardBatch(t *testing.T) {
	input := wordTableGuardFixture(t)
	cases := []struct{ name, old, new string }{
		{"hyperlink", `<w:r><w:t>Item</w:t></w:r>`, `<w:hyperlink w:anchor="A"><w:r><w:t>Item</w:t></w:r></w:hyperlink>`},
		{"content control", `<w:p w14:paraId="62154DCC" w14:textId="77777777" w:rsidR="00C36E4D" w:rsidRDefault="0019767B"><w:r><w:t>Item</w:t></w:r></w:p>`, `<w:sdt><w:sdtContent><w:p w14:paraId="62154DCC" w14:textId="77777777" w:rsidR="00C36E4D" w:rsidRDefault="0019767B"><w:r><w:t>Item</w:t></w:r></w:p></w:sdtContent></w:sdt>`},
		{"field", `<w:r><w:t>Item</w:t></w:r>`, `<w:fldSimple w:instr="PAGE"><w:r><w:t>Item</w:t></w:r></w:fldSimple>`},
		{"nested table", `<w:t>Item</w:t></w:r></w:p>`, `<w:t>Item</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p/></w:tc></w:tr></w:tbl>`},
		{"merged neighbour", `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:tcBorders>`, `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:gridSpan w:val="2"/><w:tcBorders>`},
		{"foreign cell", `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:tcBorders>`, `<w:tc other="x"><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:tcBorders>`},
		{"grid mismatch", `<w:tblGrid><w:gridCol w:w="2880"/>`, `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/>`},
		{"unknown grid child", `<w:tblGrid><w:gridCol w:w="2880"/>`, `<w:tblGrid><w:gridCol w:w="2880"/><w:bad/>`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			mutated := wordTableGuardChanged(t, input, c.old, c.new)
			s, e := OpenEditing(mutated, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target, e := s.FindRetainedTable(0, 0, 0)
			if e == nil {
				e = s.SetRetainedTable(target, map[string]any{"shading": "A1B2C3"})
			}
			var refusal *packaging.Refusal
			if !errors.As(e, &refusal) {
				t.Fatalf("not refused: %v", e)
			}
			dest := filepath.Join(t.TempDir(), "refused.docx")
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
