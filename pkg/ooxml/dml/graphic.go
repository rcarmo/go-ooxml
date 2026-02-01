package dml

import "encoding/xml"

// Graphic represents a DrawingML graphic container.
type Graphic struct {
	XMLName     xml.Name     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
	GraphicData *GraphicData `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphicData"`
}

// GraphicData holds the graphic data payload.
type GraphicData struct {
	URI string `xml:"uri,attr,omitempty"`
	Tbl *Tbl   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tbl,omitempty"`
}

// GraphicDataURITable is the URI for table data.
const GraphicDataURITable = "http://schemas.openxmlformats.org/drawingml/2006/table"
