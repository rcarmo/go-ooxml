package dml

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalTxBody(f *testing.F) {
	seeds := []string{
		`<a:txBody xmlns:a="` + NS + `"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></a:txBody>`,
		`<a:txBody xmlns:a="` + NS + `"><a:bodyPr wrap="square"/><a:p><a:r><a:t>Line</a:t></a:r></a:p></a:txBody>`,
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<a:txBody") || !strings.Contains(xmlInput, "xmlns:a") {
			return
		}

		var body TxBody
		if err := utils.UnmarshalXML([]byte(xmlInput), &body); err != nil {
			return
		}

		roundTrip, err := utils.MarshalXMLWithHeader(&body)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(roundTrip), "<txBody") || !strings.Contains(string(roundTrip), NS) {
			t.Fatalf("round-trip missing txBody element")
		}
	})
}
