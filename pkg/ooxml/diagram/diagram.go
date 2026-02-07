// Package diagram provides DrawingML diagram (SmartArt) types.
package diagram

import "encoding/xml"

// Namespaces used in DrawingML diagram parts.
const (
	NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
)

// DataModel represents a diagram data part.
type DataModel struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/diagram dataModel"`
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
