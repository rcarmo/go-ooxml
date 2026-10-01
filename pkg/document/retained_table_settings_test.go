package document

import (
	"archive/zip"
	"bytes"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"io"
	"os"
	"path/filepath"
	"testing"
)

func wordTableMutateMember(t *testing.T, input []byte, part, old, new string) []byte {
	t.Helper()
	reader, e := zip.NewReader(bytes.NewReader(input), int64(len(input)))
	if e != nil {
		t.Fatal(e)
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	changedCount := 0
	for _, f := range reader.File {
		rc, er := f.Open()
		if er != nil {
			t.Fatal(er)
		}
		data, er := io.ReadAll(rc)
		if er != nil {
			t.Fatal(er)
		}
		if er = rc.Close(); er != nil {
			t.Fatal(er)
		}
		if f.Name == part {
			if bytes.Count(data, []byte(old)) != 1 {
				t.Fatalf("member %s seed count %d", part, bytes.Count(data, []byte(old)))
			}
			data = bytes.Replace(data, []byte(old), []byte(new), 1)
			if part == "word/settings.xml" && new == `<w:notSettings ` {
				if bytes.Count(data, []byte(`</w:settings>`)) != 1 {
					t.Fatal("settings closing tag")
				}
				data = bytes.Replace(data, []byte(`</w:settings>`), []byte(`</w:notSettings>`), 1)
			}
			changedCount++
		}
		h := f.FileHeader
		h.CRC32 = 0
		h.CompressedSize = 0
		h.CompressedSize64 = 0
		h.UncompressedSize = 0
		h.UncompressedSize64 = 0
		h.Extra = nil
		w, er := zw.CreateHeader(&h)
		if er != nil {
			t.Fatal(er)
		}
		if _, er = w.Write(data); er != nil {
			t.Fatal(er)
		}
	}
	if changedCount != 1 {
		t.Fatalf("part %s count %d", part, changedCount)
	}
	if e = zw.Close(); e != nil {
		t.Fatal(e)
	}
	return buf.Bytes()
}
func TestRetainedTableWordSettingsBatch(t *testing.T) {
	input := wordTableGuardFixture(t)
	const rel = "word/_rels/document.xml.rels"
	const ct = packaging.ContentTypesPath
	const settings = "word/settings.xml"
	cases := []struct{ name, part, old, new string }{
		{"duplicate settings edge", rel, `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`, `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId77" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`},
		{"external settings edge", rel, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"`, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="https://example.invalid/settings.xml" TargetMode="External"`},
		{"settings wrong MIME", ct, `PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"`, `PartName="/word/settings.xml" ContentType="application/xml"`},
		{"settings wrong root", settings, `<w:settings `, `<w:notSettings `},
		{"settings track revisions", settings, `</w:settings>`, `<w:trackRevisions/></w:settings>`},
		{"settings protection", settings, `</w:settings>`, `<w:documentProtection w:enforcement="1"/></w:settings>`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			mutated := wordTableMutateMember(t, input, c.part, c.old, c.new)
			s, e := OpenEditing(mutated, packaging.Limits{})
			if e != nil {
				var r *packaging.Refusal
				if errors.As(e, &r) {
					return
				}
				t.Fatal(e)
			}
			target, e := s.FindRetainedTable(0, 0, 0)
			if e == nil {
				e = s.SetRetainedTable(target, map[string]any{"shading": "A1B2C3"})
			}
			var refusal *packaging.Refusal
			if !errors.As(e, &refusal) {
				t.Fatalf("not refused %v", e)
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
