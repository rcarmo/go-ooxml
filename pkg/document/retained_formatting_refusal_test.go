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

func retainedFixtureSession(t *testing.T) ([]byte, *EditSession) {
	t.Helper()
	data, e := os.ReadFile(testutil.ReferencePath("fixtures", "docx", "direct-formatting", "direct-formatting-2d32cedb722e.docx"))
	if os.IsNotExist(e) {
		t.Skip("candidate absent")
	}
	if e != nil {
		t.Fatal(e)
	}
	s, e := OpenEditing(data, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	return data, s
}
func retainedReplaceForTest(t *testing.T, s *EditSession, part, old, replacement string) {
	t.Helper()
	data, hash, e := s.pkg.Part(part)
	if e != nil {
		t.Fatal(e)
	}
	if bytes.Count(data, []byte(old)) != 1 {
		t.Fatalf("ambiguous test profile %s", old)
	}
	out := bytes.Replace(data, []byte(old), []byte(replacement), 1)
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: out}}); e != nil {
		t.Fatal(e)
	}
}
func retainedSaveForTest(t *testing.T, s *EditSession) []byte {
	t.Helper()
	dest := filepath.Join(t.TempDir(), "out.docx")
	if _, e := s.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	b, e := os.ReadFile(dest)
	if e != nil {
		t.Fatal(e)
	}
	return b
}
func retainedRefusalKind(t *testing.T, e error, kind string) {
	t.Helper()
	var r *packaging.Refusal
	if !errors.As(e, &r) || r.Kind != kind {
		t.Fatalf("refusal=%v want %s", e, kind)
	}
}
func TestRetainedFormatAtomicRefusals(t *testing.T) {
	for _, c := range []struct {
		name  string
		run   int
		patch map[string]any
		kind  string
	}{
		{"late invalid size", 0, map[string]any{"color": "A1B2C3", "sizeHalfPoints": int64(0)}, "invalid-format"},
		{"invalid enum", 0, map[string]any{"underline": "unknown"}, "invalid-format"},
		{"invalid paired font", 0, map[string]any{"ascii": "Aptos"}, "invalid-format"},
		{"invalid paragraph spacing", -1, map[string]any{"before": int64(240), "line": int64(0)}, "invalid-format"},
	} {
		t.Run(c.name, func(t *testing.T) {
			source, s := retainedFixtureSession(t)
			original := bytes.Clone(source)
			held, e := s.FindRetainedFormat(0, c.run)
			if e != nil {
				t.Fatal(e)
			}
			before := retainedSaveForTest(t, s)
			retainedRefusalKind(t, s.SetRetainedFormat(held, c.patch), c.kind)
			if e = s.retainedCurrent(held); e != nil {
				t.Fatalf("refusal lost held target: %v", e)
			}
			if !bytes.Equal(before, retainedSaveForTest(t, s)) || !bytes.Equal(source, original) {
				t.Fatal("refusal mutated archive")
			}
			valid := map[string]any{"color": "112233"}
			if c.run < 0 {
				valid = map[string]any{"outlineLvl": int64(9)}
			}
			if e = s.SetRetainedFormat(held, valid); e != nil {
				t.Fatalf("held no-op unusable: %v", e)
			}
			if !bytes.Equal(before, retainedSaveForTest(t, s)) {
				t.Fatal("held no-op mutated archive")
			}
		})
	}
}
func TestRetainedFormatTopologyAndDelivery(t *testing.T) {
	cases := []struct {
		name, part, old, new, kind string
		run                        int
	}{
		{"protected", "word/settings.xml", `</w:settings>`, `<w:documentProtection w:enforcement="1"/></w:settings>`, "protected_operation", 0},
		{"tracked", "word/settings.xml", `</w:settings>`, `<w:trackRevisions/></w:settings>`, "unsupported_structure", 0},
		{"run comment", "word/document.xml", `<w:rPr><w:rFonts`, `<w:rPr><!--blocked--><w:rFonts`, "unsupported_structure", 0},
		{"foreign leaf", "word/document.xml", `<w:color w:val="112233"/>`, `<w:color w:val="112233" xml:lang="bad"/>`, "unsupported_structure", 0},
		{"paragraph wrapper", "word/document.xml", `<w:r w:rsidR="00AB1234">`, `<w:hyperlink/><w:r w:rsidR="00AB1234">`, "unsupported_structure", -1},
		{"disordered property", "word/document.xml", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b w:val="1"/>`, `<w:b w:val="1"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`, "unsupported_structure", 0},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			_, s := retainedFixtureSession(t)
			retainedReplaceForTest(t, s, c.part, c.old, c.new)
			before := retainedSaveForTest(t, s)
			target, e := s.FindRetainedFormat(0, c.run)
			if e == nil {
				e = s.SetRetainedFormat(target, map[string]any{"strike": true})
			}
			retainedRefusalKind(t, e, c.kind)
			if !bytes.Equal(before, retainedSaveForTest(t, s)) {
				t.Fatal("topology refusal mutated archive")
			}
		})
	}
	_, s := retainedFixtureSession(t)
	target, e := s.FindRetainedFormat(0, 0)
	if e != nil {
		t.Fatal(e)
	}
	other := *target
	other.session = nil
	retainedRefusalKind(t, s.SetRetainedFormat(&other, map[string]any{"strike": true}), "stale_target")
	if e = s.SetRetainedFormat(target, map[string]any{"strike": true}); e != nil {
		t.Fatal(e)
	}
	retainedRefusalKind(t, s.SetRetainedFormat(target, map[string]any{"vanish": true}), "stale_target")
	dest := filepath.Join(t.TempDir(), "existing.docx")
	old := []byte("old destination")
	if e = os.WriteFile(dest, old, 0600); e != nil {
		t.Fatal(e)
	}
	if _, e = s.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	absent := filepath.Join(t.TempDir(), "absent.docx")
	if e = os.Mkdir(absent, 0700); e != nil {
		t.Fatal(e)
	}
	if _, e = s.SaveAs(absent); e == nil {
		t.Fatal("directory save unexpectedly succeeded")
	}
	if info, e := os.Stat(absent); e != nil || !info.IsDir() {
		t.Fatal("failed save replaced directory")
	}
}
