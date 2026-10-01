package acceptance

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func tableCorruptionFixtureAvailable(t *testing.T) {
	t.Helper()
	if _, err := os.Stat(testutil.ReferencePath("ledgers", "retained-table-properties.json")); err != nil {
		if os.IsNotExist(err) && os.Getenv("OOXML_FIXTURES_ROOT") == "" && os.Getenv("OOXML_REFERENCE_PIN") == "" {
			t.Skip("candidate absent from default pin")
		}
		t.Fatal(err)
	}
}

func TestRetainedTableOracleUnpatchedLiteralCorruption(t *testing.T) {
	tableCorruptionFixtureAvailable(t)
	records, err := tableRecords()
	if err != nil {
		t.Fatal(err)
	}
	cases := []struct{ name, id, fixture, part, owner, from, to string }{
		{"pptx border dash quote", "@id-pptx-table-properties-border-top", "fixture-ce68a5cbf3d3c25053ecc9b327b6776b35267f5b134346366e4b351ffda66abf", "ppt/slides/slide1.xml", "lnT", `val="solid"`, `val='solid'`},
		{"pptx nofill unchanged width quote", "@id-pptx-table-properties-border-no-fill", "fixture-ce68a5cbf3d3c25053ecc9b327b6776b35267f5b134346366e4b351ffda66abf", "ppt/slides/slide1.xml", "lnT", `w="12700"`, `w='12700'`},
		{"word margin dxa quote", "@id-docx-table-properties-cell-margins", "fixture-ea4dfb625197cf2ca259472419933c5c164ca733ccd765c48410d5547e0ab14f", "word/document.xml", "top", `w:type="dxa"`, `w:type='dxa'`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			path, e := testutil.LookupFixture(c.fixture)
			if e != nil {
				t.Fatal(e)
			}
			if !filepath.IsAbs(path) {
				t.Fatal("fixture not resolved")
			}
			archive, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			members, e := pptxMembers(archive)
			if e != nil {
				t.Fatal(e)
			}
			before := members[c.part]
			v, e := tableViewOf(before, records[c.id].Format)
			if e != nil {
				t.Fatal(e)
			}
			owner := v.tc
			ns := pptxA
			if records[c.id].Format == "docx" {
				ns = retainedDocxW
			}
			if c.owner == "top" {
				owner, e = tableOnly(v.doc, owner, ns, "tcMar")
				if e != nil {
					t.Fatal(e)
				}
			}
			chosen, e := tableOnly(v.doc, owner, ns, c.owner)
			if e != nil {
				t.Fatal(e)
			}
			a, b := chosen.SourceRange()
			raw := before[a:b]
			if bytes.Count(raw, []byte(c.from)) != 1 {
				t.Fatalf("selected %s lacks unique %q", c.owner, c.from)
			}
			after := append(append(bytes.Clone(before[:a]), bytes.Replace(raw, []byte(c.from), []byte(c.to), 1)...), before[b:]...)
			if e = tableMaskedSame(before, after, records[c.id]); e == nil {
				t.Fatal("unpatched literal corruption was masked")
			}
		})
	}
}
