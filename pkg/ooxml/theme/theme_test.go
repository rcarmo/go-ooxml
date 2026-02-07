package theme

import (
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestThemeRoundTrip(t *testing.T) {
	th := &Theme{
		Name:          "Office",
		ThemeElements: &ThemeElements{ClrScheme: &ClrScheme{Name: "Office"}},
	}

	data, err := utils.MarshalXMLWithHeader(th)
	if err != nil {
		t.Fatalf("MarshalXMLWithHeader error: %v", err)
	}

	var parsed Theme
	if err := utils.UnmarshalXML(data, &parsed); err != nil {
		t.Fatalf("UnmarshalXML error: %v", err)
	}
	if parsed.Name != "Office" {
		t.Fatalf("expected theme name to round-trip")
	}
	if parsed.ThemeElements == nil || parsed.ThemeElements.ClrScheme == nil {
		t.Fatalf("expected theme elements after round-trip")
	}
}
