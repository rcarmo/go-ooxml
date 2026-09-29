package losslessxml

import (
	"bytes"
	"encoding/xml"
	"testing"
)

func TestUnicodeSourceRangeControls(t *testing.T) {
	for _, tc := range []struct {
		name, source, child string
	}{
		{"Unicode prefix and attribute", `<r xmlns:名="urn:one"><名:項 名:鍵="值"/></r>`, `<名:項 名:鍵="值"/>`},
		{"Unicode alias and content", `<r xmlns:字="urn:one"><字:項 字:鍵="値">漢</字:項></r>`, `<字:項 字:鍵="値">漢</字:項>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := []byte(tc.source)
			original := bytes.Clone(source)
			doc, err := Parse(source)
			if err != nil {
				t.Fatal(err)
			}
			children := doc.Elements()
			if len(children) != 2 || children[1].Name() != (xml.Name{Space: "urn:one", Local: "項"}) {
				t.Fatalf("expanded Unicode child: %+v", children)
			}
			attrs := children[1].Attributes()
			if len(attrs) != 1 || attrs[0].Name != (xml.Name{Space: "urn:one", Local: "鍵"}) {
				t.Fatalf("expanded Unicode attribute: %+v", attrs)
			}
			start := bytes.Index(original, []byte(tc.child))
			end := start + len([]byte(tc.child))
			gotStart, gotEnd := children[1].SourceRange()
			if start < 0 || gotStart != start || gotEnd != end || !bytes.Equal(children[1].Raw(), original[start:end]) {
				t.Fatalf("Unicode byte range [%d:%d), want [%d:%d)", gotStart, gotEnd, start, end)
			}
			// The document owns its input snapshot; the caller may change its copy.
			source[start] = '!'
			if !bytes.Equal(children[1].Raw(), original[start:end]) {
				t.Fatal("source range changed with caller bytes")
			}
		})
	}
	var invalid Element
	if start, end := invalid.SourceRange(); start != -1 || end != -1 {
		t.Fatalf("invalid element range [%d:%d)", start, end)
	}
	bad := []byte(`<r><名:項/></r>`)
	original := bytes.Clone(bad)
	if doc, err := Parse(bad); err == nil || doc != nil || !bytes.Equal(bad, original) {
		t.Fatal("unbound Unicode prefix accepted or caller input changed")
	}
}
