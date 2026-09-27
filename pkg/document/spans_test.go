package document

import (
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestExactSpanMappingBatch(t *testing.T) {
	cases := []struct {
		body, needle string
		count        int
	}{{`<w:p><w:r><w:t>Al</w:t></w:r><w:r><w:t>pha</w:t></w:r></w:p>`, "Alpha", 1}, {`<w:p><w:r><w:t>😀é</w:t></w:r><w:r><w:t>Z</w:t></w:r></w:p>`, "éZ", 1}, {`<w:p><w:r><w:t>aaa</w:t></w:r></w:p>`, "aa", 1}, {`<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p>`, "A\nB", 1}, {`<w:p><w:del><w:r><w:delText>old</w:delText></w:r></w:del><w:ins><w:r><w:t>new</w:t></w:r></w:ins></w:p>`, "old", 0}, {`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><w:t>y</w:t></w:r></w:p>`, "xy", 0}}
	for _, tc := range cases {
		t.Run(tc.needle, func(t *testing.T) {
			d, err := losslessxml.Parse([]byte(`<w:document xmlns:w="` + packaging.NSWordprocessingML + `"><w:body>` + tc.body + `</w:body></w:document>`))
			if err != nil {
				t.Fatal(err)
			}
			s := &EditSession{}
			matches := s.findExact(d, "hash", tc.needle)
			if len(matches) != tc.count {
				t.Fatalf("got %d", len(matches))
			}
			for _, m := range matches {
				var got []rune
				for _, seg := range m.segments {
					r := []rune(seg.text)
					got = append(got, r[seg.start:seg.end]...)
				}
				if string(got) != tc.needle {
					t.Fatalf("mapping %q", got)
				}
			}
		})
	}
}
