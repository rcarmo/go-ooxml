package document

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestRetainedFormatNative(t *testing.T) {
	path := testutil.ReferencePath("fixtures", "docx", "direct-formatting", "direct-formatting-2d32cedb722e.docx")
	source, e := os.ReadFile(path)
	if os.IsNotExist(e) {
		t.Skip("sealed candidate fixture absent")
	}
	if e != nil {
		t.Fatal(e)
	}
	cases := []struct {
		name  string
		run   int
		patch map[string]any
		token []byte
	}{
		{"strike", 0, map[string]any{"strike": true}, []byte(`<w:strike w:val="1"/>`)},
		{"underline", 0, map[string]any{"underline": "double"}, []byte(`<w:u w:val="double"/>`)},
		{"font", 0, map[string]any{"ascii": "Aptos", "hAnsi": "Aptos"}, []byte(`<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>`)},
		{"size", 0, map[string]any{"sizeHalfPoints": int64(21)}, []byte(`<w:sz w:val="21"/>`)},
		{"inherit", 0, map[string]any{"removeDirect": true}, []byte(`<w:b w:val="1"/>`)},
		{"paragraph", -1, map[string]any{"before": int64(240), "after": int64(120)}, []byte(`w:before="240"`)},
		{"indent", -1, map[string]any{"left": int64(720), "hanging": int64(360)}, []byte(`w:hanging="360"`)},
		{"pageBreak", -1, map[string]any{"pageBreakBefore": false}, []byte(`<w:pageBreakBefore w:val="0"/>`)},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			original := bytes.Clone(source)
			s, e := OpenEditing(source, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target, e := s.FindRetainedFormat(0, c.run)
			if e != nil {
				t.Fatal(e)
			}
			if e = s.SetRetainedFormat(target, c.patch); e != nil {
				t.Fatal(e)
			}
			path := filepath.Join(t.TempDir(), "out.docx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			if bytes.Equal(saved, source) {
				t.Fatal("no edit")
			}
			if !bytes.Equal(source, original) {
				t.Fatal("actual caller buffer changed")
			}
			reopened, e := OpenEditing(saved, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			part, _, e := reopened.pkg.Part(reopened.part)
			if e != nil {
				t.Fatal(e)
			}
			if !bytes.Contains(part, c.token) {
				t.Fatalf("missing reopened %s", c.token)
			}
		})
	}
}
