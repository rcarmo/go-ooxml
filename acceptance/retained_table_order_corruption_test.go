package acceptance

import (
	"bytes"
	"os"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// Independent saved-output order oracle must reject schema-order corruption.
func TestRetainedTableOrderCorruption(t *testing.T) {
	tableCorruptionFixtureAvailable(t)
	records, e := tableRecords()
	if e != nil {
		t.Fatal(e)
	}
	cases := []struct{ name, id, fixture, part, old, new string }{
		{"pptx lines", "@id-pptx-table-properties-border-left", "fixture-ce68a5cbf3d3c25053ecc9b327b6776b35267f5b134346366e4b351ffda66abf", "ppt/slides/slide1.xml", `<a:lnL w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="solid"/></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="solid"/></a:lnR>`, `<a:lnR w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="solid"/></a:lnR><a:lnL w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="solid"/></a:lnL>`},
		{"word row flags", "@id-docx-table-properties-row-header", "fixture-ea4dfb625197cf2ca259472419933c5c164ca733ccd765c48410d5547e0ab14f", "word/document.xml", `<w:cantSplit w:val="0"/><w:trHeight w:val="240" w:hRule="atLeast"/><w:tblHeader w:val="0"/>`, `<w:tblHeader w:val="0"/><w:trHeight w:val="240" w:hRule="atLeast"/><w:cantSplit w:val="0"/>`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			path, er := testutil.LookupFixture(c.fixture)
			if er != nil {
				t.Fatal(er)
			}
			members, er := pptxMembers(mustReadTableFixture(t, path))
			if er != nil {
				t.Fatal(er)
			}
			before := members[c.part]
			if bytes.Count(before, []byte(c.old)) != 1 {
				t.Fatalf("order seed count %d", bytes.Count(before, []byte(c.old)))
			}
			after := bytes.Replace(before, []byte(c.old), []byte(c.new), 1)
			state := &tableState{record: records[c.id], after: map[string][]byte{c.part: after}, saved: []byte("candidate")}
			if er = tableOrder(state); er == nil {
				t.Fatal("order corruption escaped oracle")
			}
		})
	}
}
func mustReadTableFixture(t *testing.T, path string) []byte {
	t.Helper()
	b, e := os.ReadFile(path)
	if e != nil {
		t.Fatal(e)
	}
	return b
}
