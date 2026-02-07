package theme

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalTheme(f *testing.F) {
	seed := `<a:theme xmlns:a="` + NS + `" name="Office"><a:themeElements/></a:theme>`
	f.Add(seed)

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<a:theme") || !strings.Contains(xmlInput, "xmlns:a") {
			return
		}
		var th Theme
		if err := utils.UnmarshalXML([]byte(xmlInput), &th); err != nil {
			return
		}
		out, err := utils.MarshalXMLWithHeader(&th)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(out), "theme") {
			t.Fatalf("round-trip missing theme")
		}
	})
}
