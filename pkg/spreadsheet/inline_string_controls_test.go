package spreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const inlineWrapFixture = "fixture-38c2ed936696179d3b2359e9107ad2b8d62d71d69296f8f60bdfe1fe8f7f2439"

func inlineWrapSource(t *testing.T) []byte {
	t.Helper()
	path, err := testutil.LookupFixture(inlineWrapFixture)
	if err != nil {
		t.Fatal(err)
	}
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return b
}

// Read independent ZIP payloads rather than asking the editing session for its own result.
func inlineWrapMembers(t *testing.T, b []byte) map[string][]byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		t.Fatal(err)
	}
	out := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		if _, exists := out[f.Name]; exists {
			t.Fatalf("duplicate ZIP member %s", f.Name)
		}
		r, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, readErr := io.ReadAll(r)
		closeErr := r.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		out[f.Name] = data
	}
	return out
}

func inlineWrapPatch(t *testing.T, source []byte, part, old, replacement string) []byte {
	t.Helper()
	p, err := packaging.OpenPreserved(source, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	body, hash, err := p.Part(part)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Count(body, []byte(old)) != 1 {
		t.Fatalf("expected unique mutation anchor in %s", part)
	}
	body = bytes.Replace(body, []byte(old), []byte(replacement), 1)
	if err := p.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: body}}); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "patched.xlsx")
	if _, err := p.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	output, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return output
}

// Registry bytes cannot be replaced through Preserved.Replace; build malformed
// relationship fixtures outside production APIs.
func inlineWrapRelsPatch(t *testing.T, source []byte, old, replacement string) []byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var output bytes.Buffer
	zw := zip.NewWriter(&output)
	found := false
	for _, f := range zr.File {
		r, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, readErr := io.ReadAll(r)
		closeErr := r.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		if f.Name == "xl/_rels/workbook.xml.rels" {
			if bytes.Count(data, []byte(old)) != 1 {
				t.Fatal("relationship mutation anchor missing")
			}
			data = bytes.Replace(data, []byte(old), []byte(replacement), 1)
			found = true
		}
		h := f.FileHeader
		w, err := zw.CreateHeader(&h)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if !found {
		t.Fatal("style relationship registry missing")
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return output.Bytes()
}

func inlineWrapState(t *testing.T, s *EditSession) (sheet, styles []byte) {
	t.Helper()
	var err error
	sheet, _, err = s.pkg.Part("xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatal(err)
	}
	styles, _, err = s.pkg.Part("xl/styles.xml")
	if err != nil {
		t.Fatal(err)
	}
	return
}

func inlineWrapRefusal(t *testing.T, err error, kind string) {
	t.Helper()
	var refusal *packaging.Refusal
	if !errors.As(err, &refusal) || refusal.Kind != kind {
		t.Fatalf("expected %s refusal, got %v", kind, err)
	}
}

func TestInlineStringWrapSaveReopenControls(t *testing.T) {
	for _, tc := range []struct {
		name, value string
		wrap        bool
	}{
		{"true-special", `& < > " ' Ω`, true},
		{"false-special", "A&B <C> Ω", false},
		{"empty", "", true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := inlineWrapPatch(t, inlineWrapSource(t), "xl/worksheets/sheet1.xml", `</s:c>`, `</s:c><s:c r="B1" t="inlineStr"><s:is><s:t>Unrelated &amp; intact</s:t></s:is></s:c>`)
			original := bytes.Clone(source)
			prior := inlineWrapMembers(t, source)
			s, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if err := s.SetInlineStringWrap("Sheet", "A1", tc.value, tc.wrap); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(source, original) {
				t.Fatal("caller bytes mutated")
			}
			path := filepath.Join(t.TempDir(), "saved.xlsx")
			if _, err := s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if _, err := OpenEditing(saved, packaging.Limits{}); err != nil {
				t.Fatal(err)
			}
			after := inlineWrapMembers(t, saved)
			if len(prior) != len(after) {
				t.Fatal("member inventory changed")
			}
			for part, before := range prior {
				if part != "xl/styles.xml" && part != "xl/worksheets/sheet1.xml" && !bytes.Equal(before, after[part]) {
					t.Fatalf("unrelated member changed: %s", part)
				}
			}
			if !bytes.Contains(after["xl/styles.xml"], []byte(`count="2"`)) {
				t.Fatal("cellXfs count not incremented")
			}
			if !bytes.Contains(after["xl/worksheets/sheet1.xml"], []byte(`s="1"`)) {
				t.Fatal("cell style index not set")
			}
			const unrelatedCell = `<s:c r="B1" t="inlineStr"><s:is><s:t>Unrelated &amp; intact</s:t></s:is></s:c>`
			if !bytes.Contains(prior["xl/worksheets/sheet1.xml"], []byte(unrelatedCell)) || !bytes.Contains(after["xl/worksheets/sheet1.xml"], []byte(unrelatedCell)) {
				t.Fatal("unrelated inline cell changed")
			}
			// The new style must not rewrite the original style or unrelated style XML.
			const oldXF = `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0"/>`
			if !bytes.Contains(prior["xl/styles.xml"], []byte(oldXF)) || !bytes.Contains(after["xl/styles.xml"], []byte(oldXF)) {
				t.Fatal("original XF changed")
			}
			var xfs, wrapAttrs int
			dec := xml.NewDecoder(bytes.NewReader(after["xl/styles.xml"]))
			inCellXfs := false
			for {
				token, err := dec.Token()
				if err == io.EOF {
					break
				}
				if err != nil {
					t.Fatal(err)
				}
				switch v := token.(type) {
				case xml.StartElement:
					if v.Name.Local == "cellXfs" {
						inCellXfs = true
						if xmlAttr(v.Attr, "count") != "2" {
							t.Fatal("wrong XF count")
						}
					}
					if !inCellXfs {
						continue
					}
					if v.Name.Local == "xf" {
						xfs++
						if xfs == 2 && xmlAttr(v.Attr, "applyAlignment") != "1" {
							t.Fatal("new XF alignment not applied")
						}
					}
					if v.Name.Local == "alignment" {
						if xfs != 2 || xmlAttr(v.Attr, "wrapText") != map[bool]string{true: "true", false: "false"}[tc.wrap] {
							t.Fatal("new XF wrap incorrect")
						}
						wrapAttrs++
					}
				case xml.EndElement:
					if v.Name.Local == "cellXfs" {
						inCellXfs = false
					}
				}
			}
			if xfs != 2 || wrapAttrs != 1 {
				t.Fatalf("cell XFs=%d wrap alignments=%d", xfs, wrapAttrs)
			}
			dec = xml.NewDecoder(bytes.NewReader(after["xl/worksheets/sheet1.xml"]))
			inA1, inText, textSeen := false, false, false
			var text string
			for {
				token, err := dec.Token()
				if err == io.EOF {
					break
				}
				if err != nil {
					t.Fatal(err)
				}
				switch v := token.(type) {
				case xml.StartElement:
					if v.Name.Local == "c" && xmlAttr(v.Attr, "r") == "A1" {
						inA1 = true
						if xmlAttr(v.Attr, "s") != "1" {
							t.Fatal("wrong A1 style")
						}
					}
					if inA1 && v.Name.Local == "t" {
						inText, textSeen = true, true
					}
				case xml.CharData:
					if inText {
						text += string(v)
					}
				case xml.EndElement:
					if v.Name.Local == "t" {
						inText = false
					}
					if v.Name.Local == "c" {
						inA1 = false
					}
				}
			}
			if !textSeen || text != tc.value {
				t.Fatalf("reopened inline text = %q, want %q", text, tc.value)
			}
		})
	}
}

func xmlAttr(attrs []xml.Attr, name string) string {
	for _, a := range attrs {
		if a.Name.Space == "" && a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}

func TestInlineStringWrapRefusalCustodyControls(t *testing.T) {
	base := inlineWrapSource(t)
	const sheet = "xl/worksheets/sheet1.xml"
	const styles = "xl/styles.xml"
	for _, tc := range []struct {
		name   string
		mutate func(*testing.T, []byte) []byte
		value  string
	}{
		{"duplicate-cell", func(t *testing.T, b []byte) []byte {
			return inlineWrapPatch(t, b, sheet, `</s:c>`, `</s:c><s:c r="A1" t="inlineStr"><s:is><s:t>other</s:t></s:is></s:c>`)
		}, "new"},
		{"duplicate-style-table", func(t *testing.T, b []byte) []byte {
			return inlineWrapPatch(t, b, styles, `</cellXfs>`, `</cellXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`)
		}, "new"},
		{"wrong-style-relationship-URI", func(t *testing.T, b []byte) []byte {
			return inlineWrapRelsPatch(t, b, packaging.NSDocumentRelationships+`/styles`, packaging.NSRelationships+`/styles`)
		}, "new"},
		{"duplicate-style-relationship", func(t *testing.T, b []byte) []byte {
			return inlineWrapRelsPatch(t, b, `</Relationships>`, `<Relationship Type="`+packaging.NSDocumentRelationships+`/styles" Target="styles.xml" Id="rId4"/></Relationships>`)
		}, "new"},
		{"invalid-style-count", func(t *testing.T, b []byte) []byte {
			return inlineWrapPatch(t, b, styles, `<cellXfs count="1">`, `<cellXfs count="2">`)
		}, "new"},
		{"invalid-XML-NUL", func(t *testing.T, b []byte) []byte { return b }, "bad\x00value"},
		{"invalid-XML-control", func(t *testing.T, b []byte) []byte { return b }, "bad\x01value"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := tc.mutate(t, base)
			caller := bytes.Clone(source)
			s, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			beforeSheet, beforeStyles := inlineWrapState(t, s)
			err = s.SetInlineStringWrap("Sheet", "A1", tc.value, true)
			inlineWrapRefusal(t, err, "unsupported_structure")
			afterSheet, afterStyles := inlineWrapState(t, s)
			if !bytes.Equal(beforeSheet, afterSheet) || !bytes.Equal(beforeStyles, afterStyles) || !bytes.Equal(caller, source) {
				t.Fatal("refusal mutated parts or caller bytes")
			}
			path := filepath.Join(t.TempDir(), "refused.xlsx")
			if _, err := s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			prior, after := inlineWrapMembers(t, caller), inlineWrapMembers(t, saved)
			for name, data := range prior {
				if !bytes.Equal(data, after[name]) {
					t.Fatalf("refusal changed saved part %s", name)
				}
			}
		})
	}
}

func TestInlineStringWrapStaleTwoPartPlan(t *testing.T) {
	source := inlineWrapSource(t)
	s, err := OpenEditing(source, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	sheet, sheetHash, err := s.pkg.Part("xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatal(err)
	}
	styles, stylesHash, err := s.pkg.Part("xl/styles.xml")
	if err != nil {
		t.Fatal(err)
	}
	// A stale expected hash in the second of two parts must not stage the first.
	_, err = s.pkg.PlanGraphMutation(packaging.GraphMutation{Replacements: []packaging.Replacement{
		{Part: "xl/worksheets/sheet1.xml", ExpectedSHA256: sheetHash, Data: bytes.Replace(sheet, []byte("before"), []byte("after"), 1)},
		{Part: "xl/styles.xml", ExpectedSHA256: strings.Repeat("0", len(stylesHash)), Data: styles},
	}})
	inlineWrapRefusal(t, err, "stale_target")
	beforeSheet, beforeStyles := inlineWrapState(t, s)
	if !bytes.Equal(beforeSheet, sheet) || !bytes.Equal(beforeStyles, styles) {
		t.Fatal("preflight changed a part")
	}
	plan, err := s.pkg.PlanGraphMutation(packaging.GraphMutation{Replacements: []packaging.Replacement{
		{Part: "xl/worksheets/sheet1.xml", ExpectedSHA256: sheetHash, Data: bytes.Replace(sheet, []byte("before"), []byte("after"), 1)},
		{Part: "xl/styles.xml", ExpectedSHA256: stylesHash, Data: bytes.Replace(styles, []byte(`count="1"`), []byte(`count="2"`), 1)},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if err := s.SetInlineStringWrap("Sheet", "A1", "committed", true); err != nil {
		t.Fatal(err)
	}
	beforeSheet, beforeStyles = inlineWrapState(t, s)
	inlineWrapRefusal(t, s.pkg.ApplyGraphPlan(plan), "stale_target")
	afterSheet, afterStyles := inlineWrapState(t, s)
	if !bytes.Equal(beforeSheet, afterSheet) || !bytes.Equal(beforeStyles, afterStyles) || !bytes.Equal(source, inlineWrapSource(t)) {
		t.Fatal("stale plan mutated parts or caller")
	}
	path := filepath.Join(t.TempDir(), "committed.xlsx")
	if _, err := s.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	saved, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	members := inlineWrapMembers(t, saved)
	if !bytes.Equal(members["xl/worksheets/sheet1.xml"], beforeSheet) || !bytes.Equal(members["xl/styles.xml"], beforeStyles) {
		t.Fatal("stale plan changed saved parts")
	}
}
