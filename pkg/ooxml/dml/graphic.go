package dml

import "encoding/xml"

// Graphic represents a DrawingML graphic container.
type Graphic struct {
	XMLName     xml.Name     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
	GraphicData *GraphicData `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphicData"`
}

// GraphicData holds the graphic data payload.
type GraphicData struct {
	URI     string `xml:"uri,attr,omitempty"`
	Tbl     *Tbl   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tbl,omitempty"`
	Chart   *ChartRef
	Diagram *DiagramRef
	Picture *PictureRef
	Content []interface{} `xml:",any"`
}

// GraphicDataURITable is the URI for table data.
const GraphicDataURITable = "http://schemas.openxmlformats.org/drawingml/2006/table"

// GraphicDataURIChart is the URI for chart data.
const GraphicDataURIChart = "http://schemas.openxmlformats.org/drawingml/2006/chart"
// GraphicDataURIDiagram is the URI for diagram data.
const GraphicDataURIDiagram = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
// GraphicDataURIPicture is the URI for picture data.
const GraphicDataURIPicture = "http://schemas.openxmlformats.org/drawingml/2006/picture"

// ChartRef represents a chart reference inside graphicData.
type ChartRef struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart chart"`
	RID     string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// DiagramRef represents SmartArt diagram relationship IDs.
type DiagramRef struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/diagram relIds"`
	Data    string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships dm,attr,omitempty"`
	Layout  string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships lo,attr,omitempty"`
	Colors  string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships cs,attr,omitempty"`
	Style   string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships qs,attr,omitempty"`
}

// PictureRef represents picture data inside graphicData.
type PictureRef struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/picture pic"`
	Attrs   []xml.Attr `xml:",attr"`
	Inner   string   `xml:",innerxml"`
}

// RawGraphicData preserves raw graphic data elements.
type RawGraphicData struct {
	XMLName xml.Name
	Inner   string `xml:",innerxml"`
}

// UnmarshalXML implements custom XML unmarshaling for GraphicData.
func (g *GraphicData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "uri" {
			g.URI = attr.Value
			break
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tbl":
				g.Tbl = &Tbl{}
				if err := d.DecodeElement(g.Tbl, &t); err != nil {
					return err
				}
			case "chart":
				g.Chart = &ChartRef{}
				if err := d.DecodeElement(g.Chart, &t); err != nil {
					return err
				}
		case "relIds":
			g.Diagram = &DiagramRef{}
			if err := d.DecodeElement(g.Diagram, &t); err != nil {
				return err
			}
		case "pic":
			g.Picture = &PictureRef{}
			if err := d.DecodeElement(g.Picture, &t); err != nil {
				return err
			}
		default:
			raw := &RawGraphicData{}
			if err := d.DecodeElement(raw, &t); err != nil {
				return err
			}
			g.Content = append(g.Content, raw)
			}
		case xml.EndElement:
			if t.Name == start.Name {
				return nil
			}
		}
	}
}

// MarshalXML implements custom XML marshaling for GraphicData.
func (g *GraphicData) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "graphicData"}
	if g.URI != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Local: "uri"}, Value: g.URI})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if g.Tbl != nil {
		if err := e.EncodeElement(g.Tbl, xml.StartElement{Name: xml.Name{Space: NS, Local: "tbl"}}); err != nil {
			return err
		}
	}
	if g.Chart != nil {
		if err := e.EncodeElement(g.Chart, xml.StartElement{Name: xml.Name{Space: GraphicDataURIChart, Local: "chart"}}); err != nil {
			return err
		}
	}
	if g.Diagram != nil {
		if err := e.EncodeElement(g.Diagram, xml.StartElement{Name: xml.Name{Space: GraphicDataURIDiagram, Local: "relIds"}}); err != nil {
			return err
		}
	}
	if g.Picture != nil {
		if err := e.EncodeElement(g.Picture, xml.StartElement{Name: xml.Name{Space: GraphicDataURIPicture, Local: "pic"}}); err != nil {
			return err
		}
	}
	for _, elem := range g.Content {
		if err := e.Encode(elem); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}
