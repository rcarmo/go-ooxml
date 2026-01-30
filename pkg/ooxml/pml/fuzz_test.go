package pml

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalSlide(f *testing.F) {
	seeds := []string{
		`<p:sld xmlns:p="` + NS + `" xmlns:a="` + NSA + `"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`,
		`<p:sld xmlns:p="` + NS + `" xmlns:a="` + NSA + `"><p:cSld name="Title"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`,
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<p:sld") || !strings.Contains(xmlInput, "xmlns:p") {
			return
		}

		var sld Sld
		if err := utils.UnmarshalXML([]byte(xmlInput), &sld); err != nil {
			return
		}

		roundTrip, err := utils.MarshalXMLWithHeader(&sld)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(roundTrip), "<sld") || !strings.Contains(string(roundTrip), NS) {
			t.Fatalf("round-trip missing sld element")
		}
	})
}
