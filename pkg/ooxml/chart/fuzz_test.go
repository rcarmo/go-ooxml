package chart

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalChartSpace(f *testing.F) {
	seed := `<c:chartSpace xmlns:c="` + NS + `"><c:chart/></c:chartSpace>`
	f.Add(seed)

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<c:chartSpace") || !strings.Contains(xmlInput, "xmlns:c") {
			return
		}
		var cs ChartSpace
		if err := utils.UnmarshalXML([]byte(xmlInput), &cs); err != nil {
			return
		}
		out, err := utils.MarshalXMLWithHeader(&cs)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(out), "chartSpace") {
			t.Fatalf("round-trip missing chartSpace")
		}
	})
}
