package document

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestStoryProjectionBatch(t *testing.T) {
	wrap := func(body string) *losslessxml.Document {
		t.Helper()
		d, err := losslessxml.Parse([]byte(`<w:document xmlns:w="` + packaging.NSWordprocessingML + `" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body>` + body + `</w:body></w:document>`))
		if err != nil {
			t.Fatal(err)
		}
		return d
	}
	t.Run("moves and block insertions", func(t *testing.T) {
		d := wrap(`<w:moveFrom><w:p><w:r><w:delText>old</w:delText></w:r></w:p></w:moveFrom><w:moveTo><w:p><w:r><w:t>new</w:t></w:r></w:p></w:moveTo>`)
		for _, tc := range []struct {
			v     View
			text  string
			count int
		}{{CurrentView, "new", 1}, {OriginalView, "old", 1}, {AllView, "oldnew", 2}} {
			blocks, _ := outlinePart(d, "main", tc.v)
			text := ""
			for _, b := range blocks {
				text += b.Text
			}
			if len(blocks) != tc.count || text != tc.text {
				t.Fatalf("%s %+v", tc.v, blocks)
			}
		}
	})
	t.Run("controls table and breaks", func(t *testing.T) {
		d := wrap(`<w:sdt><w:sdtContent><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>`)
		blocks, warn := outlinePart(d, "main", CurrentView)
		if len(blocks) != 1 || blocks[0].Text != "A\tB\nC" || len(warn) != 0 {
			t.Fatalf("%+v %v", blocks, warn)
		}
	})
	t.Run("alternate content and field warnings", func(t *testing.T) {
		d := wrap(`<mc:AlternateContent><mc:Choice Requires="w"><w:p><w:r><w:t>choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><w:p><w:r><w:instrText>DATE</w:instrText><w:t>cached</w:t></w:r></w:p>`)
		blocks, warn := outlinePart(d, "main", CurrentView)
		if len(blocks) != 1 || blocks[0].Text != "cached" || len(warn) != 2 || !strings.Contains(strings.Join(warn, " "), "branch selection") {
			t.Fatalf("%+v %v", blocks, warn)
		}
	})
}
