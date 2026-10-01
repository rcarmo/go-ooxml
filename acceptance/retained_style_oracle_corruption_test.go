package acceptance

import (
	"bytes"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func TestRetainedStyleOracleCorruption(t *testing.T) {
	// Independent order oracle must reject direct and nested line order changes.
	for _, c := range []struct{ name, xml, kind string }{
		{"shape fill after line", `<p:spPr xmlns:p="` + pptxP + `" xmlns:a="` + pptxA + `"><a:ln><a:noFill/></a:ln><a:solidFill><a:srgbClr val="A1B2C3"/></a:solidFill></p:spPr>`, "shape-solid-fill"},
		{"line dash before fill", `<p:spPr xmlns:p="` + pptxP + `" xmlns:a="` + pptxA + `"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln><a:prstDash val="dash"/><a:noFill/></a:ln></p:spPr>`, "shape-solid-fill"},
		{"Word vanish after color", `<w:rPr xmlns:w="` + retainedDocxW + `"><w:color w:val="112233"/><w:vanish w:val="1"/></w:rPr>`, "run-hidden"},
	} {
		t.Run(c.name, func(t *testing.T) {
			d, e := losslessxml.Parse([]byte(c.xml))
			if e != nil {
				t.Fatal(e)
			}
			ns := pptxA
			order := map[string]int{"solidFill": 1, "noFill": 1, "ln": 2}
			if c.kind == "run-hidden" {
				ns = retainedDocxW
				order = map[string]int{"vanish": 1, "color": 2}
			}
			e = retainedOrderedChildren(d, d.Elements()[0], ns, order)
			if c.name == "line dash before fill" {
				line := retainedChild(d, d.Elements()[0], pptxA, "ln")
				if len(line) != 1 {
					t.Fatal("missing nested line")
				}
				e = retainedOrderedChildren(d, line[0], pptxA, map[string]int{"solidFill": 1, "noFill": 1, "prstDash": 2})
			}
			if e == nil {
				t.Fatal("order corruption accepted")
			}
		})
	}
	before := []byte(`<w:rPr xmlns:w="` + retainedDocxW + `"><w:color w:val="112233"/><w:lang w:val="en-US"/></w:rPr>`)
	after := bytes.Replace(before, []byte(`en-US`), []byte(`pt-PT`), 1)
	a, e := retainedMaskWord(before, map[string]bool{}, map[string]bool{"color": true})
	if e != nil {
		t.Fatal(e)
	}
	b, e := retainedMaskWord(after, map[string]bool{}, map[string]bool{"color": true})
	if e != nil {
		t.Fatal(e)
	}
	if bytes.Equal(a, b) {
		t.Fatal("unpatched language corruption hidden by mask")
	}
}
