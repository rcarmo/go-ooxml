// Package diagram provides DrawingML diagram (SmartArt) types.
package diagram

import "encoding/xml"

// Namespaces used in DrawingML diagram parts.
const (
	NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
)

// DataModel represents a diagram data part.
type DataModel struct {
	XMLName  xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/diagram dataModel"`
	InnerXML string   `xml:",innerxml"`
}

// LayoutDef represents a diagram layout definition part.
type LayoutDef struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/diagram layoutDef"`
}

// StyleDef represents a diagram style definition part.
type StyleDef struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/diagram styleDef"`
}

// ColorsDef represents a diagram color definition part.
type ColorsDef struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/diagram colorsDef"`
}

// DefaultDataModel returns a minimal SmartArt data model payload.
func DefaultDataModel() *DataModel {
	return &DataModel{
		XMLName: xml.Name{Space: NS, Local: "dataModel"},
		InnerXML: `<ptLst>` +
			`<pt modelId="{00000000-0000-0000-0000-000000000001}" type="doc"><prSet/></pt>` +
			`<pt modelId="{00000000-0000-0000-0000-000000000002}" type="node"><prSet/><t>Diagram</t></pt>` +
			`</ptLst>` +
			`<cxnLst>` +
			`<cxn modelId="{00000000-0000-0000-0000-000000000003}" type="parOf" srcId="{00000000-0000-0000-0000-000000000001}" destId="{00000000-0000-0000-0000-000000000002}" srcOrd="0" destOrd="0"/>` +
			`</cxnLst>`,
	}
}

// DefaultLayoutDef returns a minimal SmartArt layout definition payload.
func DefaultLayoutDef() *LayoutDef {
	return &LayoutDef{
		XMLName: xml.Name{Space: NS, Local: "layoutDef"},
	}
}

// DefaultStyleDef returns a minimal SmartArt style definition payload.
func DefaultStyleDef() *StyleDef {
	return &StyleDef{
		XMLName: xml.Name{Space: NS, Local: "styleDef"},
	}
}

// DefaultColorsDef returns a minimal SmartArt colors definition payload.
func DefaultColorsDef() *ColorsDef {
	return &ColorsDef{
		XMLName: xml.Name{Space: NS, Local: "colorsDef"},
	}
}
