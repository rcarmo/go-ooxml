package wml

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalParagraph(f *testing.F) {
	seeds := []string{
		`<w:p xmlns:w="` + NS + `"><w:r><w:t>hello</w:t></w:r></w:p>`,
		`<w:p xmlns:w="` + NS + `"><w:r><w:t xml:space="preserve"> spaced </w:t></w:r></w:p>`,
		`<w:p xmlns:w="` + NS + `"><w:ins><w:r><w:t>ins</w:t></w:r></w:ins></w:p>`,
		`<w:p xmlns:w="` + NS + `"><w:del><w:r><w:delText>del</w:delText></w:r></w:del></w:p>`,
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<w:p") {
			return
		}
		if !strings.Contains(xmlInput, "xmlns:w") {
			return
		}

		var p P
		if err := utils.UnmarshalXML([]byte(xmlInput), &p); err != nil {
			return
		}

		roundTrip, err := utils.MarshalXMLWithHeader(&p)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(roundTrip), "<p") || !strings.Contains(string(roundTrip), NS) {
			t.Fatalf("round-trip missing paragraph element")
		}
	})
}

func FuzzUnmarshalRun(f *testing.F) {
	seeds := []string{
		`<w:r xmlns:w="` + NS + `"><w:t>text</w:t></w:r>`,
		`<w:r xmlns:w="` + NS + `"><w:t xml:space="preserve">x</w:t><w:tab/></w:r>`,
		`<w:r xmlns:w="` + NS + `"><w:br/><w:t>line</w:t></w:r>`,
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<w:r") {
			return
		}
		if !strings.Contains(xmlInput, "xmlns:w") {
			return
		}

		var r R
		if err := utils.UnmarshalXML([]byte(xmlInput), &r); err != nil {
			return
		}

		roundTrip, err := utils.MarshalXMLWithHeader(&r)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(roundTrip), "<r") || !strings.Contains(string(roundTrip), NS) {
			t.Fatalf("round-trip missing run element")
		}
	})
}

func FuzzUnmarshalBody(f *testing.F) {
	seeds := []string{
		`<w:body xmlns:w="` + NS + `"><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body>`,
		`<w:body xmlns:w="` + NS + `"><w:p/><w:tbl/></w:body>`,
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<w:body") {
			return
		}
		if !strings.Contains(xmlInput, "xmlns:w") {
			return
		}

		var body Body
		if err := utils.UnmarshalXML([]byte(xmlInput), &body); err != nil {
			return
		}

		roundTrip, err := utils.MarshalXMLWithHeader(&body)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(roundTrip), "<body") || !strings.Contains(string(roundTrip), NS) {
			t.Fatalf("round-trip missing body element")
		}
	})
}
