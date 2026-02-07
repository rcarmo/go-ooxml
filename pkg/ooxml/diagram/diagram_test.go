package diagram

import (
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestDiagramPartsRoundTrip(t *testing.T) {
	tests := []struct {
		name string
		doc  interface{}
	}{
		{"dataModel", &DataModel{}},
		{"layoutDef", &LayoutDef{}},
		{"styleDef", &StyleDef{}},
		{"colorsDef", &ColorsDef{}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			data, err := utils.MarshalXMLWithHeader(tt.doc)
			if err != nil {
				t.Fatalf("MarshalXMLWithHeader error: %v", err)
			}
			switch tt.name {
			case "dataModel":
				var parsed DataModel
				if err := utils.UnmarshalXML(data, &parsed); err != nil {
					t.Fatalf("UnmarshalXML error: %v", err)
				}
			case "layoutDef":
				var parsed LayoutDef
				if err := utils.UnmarshalXML(data, &parsed); err != nil {
					t.Fatalf("UnmarshalXML error: %v", err)
				}
			case "styleDef":
				var parsed StyleDef
				if err := utils.UnmarshalXML(data, &parsed); err != nil {
					t.Fatalf("UnmarshalXML error: %v", err)
				}
			case "colorsDef":
				var parsed ColorsDef
				if err := utils.UnmarshalXML(data, &parsed); err != nil {
					t.Fatalf("UnmarshalXML error: %v", err)
				}
			}
		})
	}
}
