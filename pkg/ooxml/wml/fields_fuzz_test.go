package wml

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalFldChar(f *testing.F) {
	seed := `<w:fldChar xmlns:w="` + NS + `" w:fldCharType="begin"/>`
	f.Add(seed)

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "fldChar") || !strings.Contains(xmlInput, "xmlns:w") {
			return
		}
		var fc FldChar
		if err := utils.UnmarshalXML([]byte(xmlInput), &fc); err != nil {
			return
		}
		out, err := utils.MarshalXMLWithHeader(&fc)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(out), "fldChar") {
			t.Fatalf("round-trip missing fldChar")
		}
	})
}
