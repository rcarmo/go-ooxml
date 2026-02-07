package dml

import "encoding/xml"

// WPInline represents a WordprocessingML inline drawing container.
type WPInline struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing inline"`
	Ext     *WPSize  `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing extent"`
	DocPr   *DocPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing docPr"`
	Graphic *Graphic `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
}

// WPAnchor represents a WordprocessingML anchored drawing container.
type WPAnchor struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing anchor"`
	Ext     *WPSize  `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing extent"`
	DocPr   *DocPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing docPr"`
	Graphic *Graphic `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
}

// WPSize represents size/extent for a WordprocessingML drawing.
type WPSize struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

// DocPr represents drawing document properties.
type DocPr struct {
	ID    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Descr string `xml:"descr,attr,omitempty"`
	Title string `xml:"title,attr,omitempty"`
}
