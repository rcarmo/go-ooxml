package diagram

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalDiagramParts(f *testing.F) {
	seeds := []string{
		`<dgm:dataModel xmlns:dgm="` + NS + `"></dgm:dataModel>`,
		`<dgm:layoutDef xmlns:dgm="` + NS + `"></dgm:layoutDef>`,
		`<dgm:styleDef xmlns:dgm="` + NS + `"></dgm:styleDef>`,
		`<dgm:colorsDef xmlns:dgm="` + NS + `"></dgm:colorsDef>`,
		`<dsp:drawing xmlns:dsp="` + NSDiagramDrawing + `" xmlns:dgm="` + NS + `"></dsp:drawing>`,
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "xmlns:dgm") {
			return
		}
		switch {
		case strings.Contains(xmlInput, "dataModel"):
			var d DataModel
			if err := utils.UnmarshalXML([]byte(xmlInput), &d); err != nil {
				return
			}
			out, err := utils.MarshalXMLWithHeader(&d)
			if err != nil {
				t.Fatalf("MarshalXMLWithHeader error: %v", err)
			}
			if !strings.Contains(string(out), "dataModel") {
				t.Fatalf("round-trip missing dataModel")
			}
		case strings.Contains(xmlInput, "layoutDef"):
			var l LayoutDef
			if err := utils.UnmarshalXML([]byte(xmlInput), &l); err != nil {
				return
			}
			out, err := utils.MarshalXMLWithHeader(&l)
			if err != nil {
				t.Fatalf("MarshalXMLWithHeader error: %v", err)
			}
			if !strings.Contains(string(out), "layoutDef") {
				t.Fatalf("round-trip missing layoutDef")
			}
		case strings.Contains(xmlInput, "styleDef"):
			var s StyleDef
			if err := utils.UnmarshalXML([]byte(xmlInput), &s); err != nil {
				return
			}
			out, err := utils.MarshalXMLWithHeader(&s)
			if err != nil {
				t.Fatalf("MarshalXMLWithHeader error: %v", err)
			}
			if !strings.Contains(string(out), "styleDef") {
				t.Fatalf("round-trip missing styleDef")
			}
		case strings.Contains(xmlInput, "colorsDef"):
			var c ColorsDef
			if err := utils.UnmarshalXML([]byte(xmlInput), &c); err != nil {
				return
			}
			out, err := utils.MarshalXMLWithHeader(&c)
			if err != nil {
				t.Fatalf("MarshalXMLWithHeader error: %v", err)
			}
			if !strings.Contains(string(out), "colorsDef") {
				t.Fatalf("round-trip missing colorsDef")
			}
		case strings.Contains(xmlInput, "drawing"):
			var d Drawing
			if err := utils.UnmarshalXML([]byte(xmlInput), &d); err != nil {
				return
			}
			out, err := utils.MarshalXMLWithHeader(&d)
			if err != nil {
				t.Fatalf("MarshalXMLWithHeader error: %v", err)
			}
			if !strings.Contains(string(out), "drawing") {
				t.Fatalf("round-trip missing drawing")
			}
		}
	})
}
