package sml

import (
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestWorksheetDrawingRoundTrip(t *testing.T) {
	ws := &Worksheet{
		Drawing:   &Drawing{ID: "rId7"},
		SheetData: &SheetData{},
	}

	data, err := utils.MarshalXMLWithHeader(ws)
	if err != nil {
		t.Fatalf("MarshalXMLWithHeader error: %v", err)
	}

	var parsed Worksheet
	if err := utils.UnmarshalXML(data, &parsed); err != nil {
		t.Fatalf("UnmarshalXML error: %v", err)
	}
	if parsed.Drawing == nil || parsed.Drawing.ID != "rId7" {
		t.Fatalf("expected drawing relationship to round-trip")
	}
}
