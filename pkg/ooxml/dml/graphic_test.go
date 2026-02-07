package dml

import (
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestGraphicDataChartRoundTrip(t *testing.T) {
	g := &Graphic{
		GraphicData: &GraphicData{
			URI:   GraphicDataURIChart,
			Chart: &ChartRef{RID: "rId5"},
		},
	}

	data, err := utils.MarshalXMLWithHeader(g)
	if err != nil {
		t.Fatalf("MarshalXMLWithHeader error: %v", err)
	}

	var parsed Graphic
	if err := utils.UnmarshalXML(data, &parsed); err != nil {
		t.Fatalf("UnmarshalXML error: %v", err)
	}
	if parsed.GraphicData == nil || parsed.GraphicData.Chart == nil {
		t.Fatalf("expected chart reference after round-trip")
	}
	if parsed.GraphicData.Chart.RID != "rId5" {
		t.Fatalf("expected chart rId to round-trip")
	}
}
