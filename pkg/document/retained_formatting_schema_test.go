package document

import (
	"bytes"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func TestRetainedFormatSchemaInsertions(t *testing.T) {
	for _, tc := range []struct {
		name    string
		source  string
		order   map[string]int
		updates map[string][]byte
		tokens  []string
	}{
		{"selfclosing run", `<w:rPr/>`, retainedRunOrder, map[string][]byte{"vanish": retainedLeaf("vanish", "1"), "strike": retainedLeaf("strike", "1")}, []string{`<w:strike w:val="1"/>`, `<w:vanish w:val="1"/>`}},
		{"selfclosing paragraph", `<w:pPr />`, retainedParaOrder, map[string][]byte{"outlineLvl": retainedLeaf("outlineLvl", "0"), "keepNext": retainedLeaf("keepNext", "1")}, []string{`<w:keepNext w:val="1"/>`, `<w:outlineLvl w:val="0"/>`}},
		{"existing ordering", `<w:rPr><w:b w:val="1"/><w:color w:val="112233"/></w:rPr>`, retainedRunOrder, map[string][]byte{"vanish": retainedLeaf("vanish", "1")}, []string{`<w:b w:val="1"/>`, `<w:vanish w:val="1"/>`, `<w:color w:val="112233"/>`}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			out, e := retainedRender([]byte(tc.source), tc.order, tc.updates)
			if e != nil {
				t.Fatal(e)
			}
			last := -1
			for _, token := range tc.tokens {
				at := bytes.Index(out, []byte(token))
				if at < 0 || at <= last {
					t.Fatalf("schema order lost: %s", out)
				}
				last = at
			}
			prefix := []byte(`<root xmlns:w="` + retainedW + `">`)
			doc, e := losslessxml.Parse(append(append(bytes.Clone(prefix), out...), []byte(`</root>`)...))
			if e != nil {
				t.Fatal(e)
			}
			if _, e = retainedOrder(doc, doc.Elements()[1], tc.order); e != nil {
				t.Fatal(e)
			}
		})
	}
	original := []byte(`<w:spacing w:before='0 > w:after="noise"' w:after="0"/>`)
	updated, e := retainedOpening(original, map[string]*string{"w:after": func() *string { x := "120"; return &x }()})
	if e != nil {
		t.Fatal(e)
	}
	if !bytes.Contains(updated, []byte(`w:before='0 > w:after="noise"'`)) || !bytes.Contains(updated, []byte(`w:after="120"`)) {
		t.Fatalf("quoted > lost: %s", updated)
	}
}
