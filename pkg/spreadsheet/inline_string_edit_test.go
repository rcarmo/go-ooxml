package spreadsheet

import (
	"bytes"
	"encoding/xml"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestInlineStringWrapRealFixture(t *testing.T) {
	fixture, err := testutil.LookupFixture("fixture-38c2ed936696179d3b2359e9107ad2b8d62d71d69296f8f60bdfe1fe8f7f2439")
	if err != nil {
		t.Fatal(err)
	}
	original, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name         string
		alias, wrong bool
	}{{name: "ordinary"}, {name: "alias", alias: true}, {name: "wrong URI", wrong: true}} {
		t.Run(tc.name, func(t *testing.T) {
			source := bytes.Clone(original)
			if tc.alias || tc.wrong {
				p, err := packaging.OpenPreserved(source, packaging.Limits{})
				if err != nil {
					t.Fatal(err)
				}
				main, hash, err := p.Part("xl/workbook.xml")
				if err != nil {
					t.Fatal(err)
				}
				if tc.alias {
					main = bytes.Replace(main, []byte(`xmlns:r="`+packaging.NSDocumentRelationships+`"`), []byte(`xmlns:link="`+packaging.NSDocumentRelationships+`"`), 1)
					main = bytes.ReplaceAll(main, []byte(`r:id=`), []byte(`link:id=`))
				} else {
					main = bytes.Replace(main, []byte(`xmlns:r="`+packaging.NSDocumentRelationships+`"`), []byte(`xmlns:r="`+packaging.NSRelationships+`"`), 1)
				}
				if err = p.Replace([]packaging.Replacement{{Part: "xl/workbook.xml", ExpectedSHA256: hash, Data: main}}); err != nil {
					t.Fatal(err)
				}
				path := filepath.Join(t.TempDir(), "alias.xlsx")
				if _, err = p.SaveAs(path); err != nil {
					t.Fatal(err)
				}
				source, err = os.ReadFile(path)
				if err != nil {
					t.Fatal(err)
				}
			}
			prior := bytes.Clone(source)
			session, err := OpenEditing(source, packaging.Limits{})
			if tc.wrong {
				var refused *packaging.Refusal
				if session != nil || !errors.As(err, &refused) || refused.Kind != "xlsx-workbook-invalid" {
					t.Fatalf("wrong URI accepted or wrong reason: session=%v error=%v", session, err)
				}
				if !bytes.Equal(source, prior) {
					t.Fatal("caller bytes changed")
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if err = session.SetInlineStringWrap("Sheet", "A1", "Edited\ncell", true); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(source, prior) {
				t.Fatal("caller bytes mutated")
			}
			output := filepath.Join(t.TempDir(), "output.xlsx")
			if _, err = session.SaveAs(output); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(output)
			if err != nil {
				t.Fatal(err)
			}
			reopened, err := OpenEditing(saved, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if _, err := reopened.pkg.Graph(); err != nil {
				t.Fatal(err)
			}
			sheet, _, err := reopened.pkg.Part("xl/worksheets/sheet1.xml")
			if err != nil {
				t.Fatal(err)
			}
			d, err := losslessxml.Parse(sheet)
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, e := range d.Elements() {
				if e.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "c"}) {
					continue
				}
				if attr(e, "r") != "A1" {
					continue
				}
				found = true
				if attr(e, "s") != "1" {
					t.Fatalf("new style index: %s", attr(e, "s"))
				}
				for _, leaf := range d.Elements() {
					owner, ok := leaf.Parent()
					if ok && owner.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "is"}) {
						value, ok := leaf.Text()
						if ok && value != "Edited\ncell" {
							t.Fatalf("inline text %q", value)
						}
					}
				}
			}
			if !found {
				t.Fatal("cell missing")
			}
			styles, _, err := reopened.pkg.Part("xl/styles.xml")
			if err != nil {
				t.Fatal(err)
			}
			sd, err := losslessxml.Parse(styles)
			if err != nil {
				t.Fatal(err)
			}
			wrap := false
			for _, e := range sd.Elements() {
				if e.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "alignment"}) && attr(e, "wrapText") == "true" {
					wrap = true
				}
			}
			if !wrap {
				t.Fatal("new wrap XF missing")
			}
			if tc.alias {
				main, _, err := reopened.pkg.Part("xl/workbook.xml")
				if err != nil {
					t.Fatal(err)
				}
				if !strings.Contains(string(main), `link:id=`) || strings.Contains(string(main), `r:id=`) {
					t.Fatalf("alias spelling not preserved: %s", main)
				}
			}
		})
	}
}
