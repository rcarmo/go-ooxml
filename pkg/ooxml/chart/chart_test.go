package chart

import (
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestChartSpaceRoundTrip(t *testing.T) {
	cs := DefaultChartSpace()

	data, err := utils.MarshalXMLWithHeader(cs)
	if err != nil {
		t.Fatalf("MarshalXMLWithHeader error: %v", err)
	}

	var parsed ChartSpace
	if err := utils.UnmarshalXML(data, &parsed); err != nil {
		t.Fatalf("UnmarshalXML error: %v", err)
	}

	if parsed.Chart == nil || parsed.Chart.PlotArea == nil {
		t.Fatalf("expected chart/plotArea after round-trip")
	}
	if parsed.Chart.Title == nil || parsed.Chart.Legend == nil {
		t.Fatalf("expected title/legend after round-trip")
	}
	if len(parsed.Chart.PlotArea.Content) == 0 {
		t.Fatalf("expected plotArea content after round-trip")
	}
}
