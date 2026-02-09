package wml

import "encoding/xml"

// Tbl represents a table.
type Tbl struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tbl"`
	TblPr   *TblPr   `xml:"tblPr,omitempty"`
	TblGrid *TblGrid `xml:"tblGrid,omitempty"`
	Tr      []*Tr    `xml:"tr,omitempty"`
}

// TblPr represents table properties.
type TblPr struct {
	TblStyle     *TblStyle     `xml:"tblStyle,omitempty"`
	TblW         *TblWidth     `xml:"tblW,omitempty"`
	Jc           *Jc           `xml:"jc,omitempty"`
	TblInd       *TblWidth     `xml:"tblInd,omitempty"`
	TblBorders   *TblBorders   `xml:"tblBorders,omitempty"`
	TblLayout    *TblLayout    `xml:"tblLayout,omitempty"`
	TblCellMar   *TblCellMar   `xml:"tblCellMar,omitempty"`
	TblLook      *TblLook      `xml:"tblLook,omitempty"`
}

// TblStyle references a table style.
type TblStyle struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// TblWidth represents table width.
type TblWidth struct {
	W    int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main w,attr"`
	Type string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main type,attr"`
}

// TblGrid represents table grid (column widths).
type TblGrid struct {
	GridCol []*GridCol `xml:"gridCol,omitempty"`
}

// GridCol represents a grid column.
type GridCol struct {
	W int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main w,attr,omitempty"`
}

// TblBorders represents table borders.
type TblBorders struct {
	Top     *Border `xml:"top,omitempty"`
	Left    *Border `xml:"left,omitempty"`
	Bottom  *Border `xml:"bottom,omitempty"`
	Right   *Border `xml:"right,omitempty"`
	InsideH *Border `xml:"insideH,omitempty"`
	InsideV *Border `xml:"insideV,omitempty"`
}

// Border represents a border.
type Border struct {
	Val   string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	Sz    int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sz,attr,omitempty"`
	Space int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main space,attr,omitempty"`
	Color string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main color,attr,omitempty"`
}

// TblLayout represents table layout mode.
type TblLayout struct {
	Type string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main type,attr"`
}

// TblCellMar represents table cell margins.
type TblCellMar struct {
	Top    *TblWidth `xml:"top,omitempty"`
	Left   *TblWidth `xml:"left,omitempty"`
	Bottom *TblWidth `xml:"bottom,omitempty"`
	Right  *TblWidth `xml:"right,omitempty"`
}

// TblLook represents table style options.
type TblLook struct {
	Val          string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	FirstRow     *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main firstRow,attr,omitempty"`
	LastRow      *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main lastRow,attr,omitempty"`
	FirstColumn  *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main firstColumn,attr,omitempty"`
	LastColumn   *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main lastColumn,attr,omitempty"`
	NoHBand      *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main noHBand,attr,omitempty"`
	NoVBand      *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main noVBand,attr,omitempty"`
}

// Tr represents a table row.
type Tr struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tr"`
	TrPr    *TrPr    `xml:"trPr,omitempty"`
	Tc      []*Tc    `xml:"tc,omitempty"`
}

// TrPr represents table row properties.
type TrPr struct {
	TrHeight *TrHeight `xml:"trHeight,omitempty"`
	TblHeader *OnOff   `xml:"tblHeader,omitempty"`
	CantSplit *OnOff   `xml:"cantSplit,omitempty"`
}

// TrHeight represents row height.
type TrHeight struct {
	Val    int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	HRule  string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hRule,attr,omitempty"`
}

// Tc represents a table cell.
type Tc struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tc"`
	TcPr    *TcPr    `xml:"tcPr,omitempty"`
	Content []interface{} `xml:",any"` // Can contain P, Tbl, etc.
}

// TcPr represents table cell properties.
type TcPr struct {
	TcW        *TblWidth      `xml:"tcW,omitempty"`
	GridSpan   *GridSpan      `xml:"gridSpan,omitempty"`
	VMerge     *VMerge        `xml:"vMerge,omitempty"`
	TcBorders  *TcBorders     `xml:"tcBorders,omitempty"`
	Shd        *Shd           `xml:"shd,omitempty"`
	VAlign     *VAlign        `xml:"vAlign,omitempty"`
	NoWrap     *OnOff         `xml:"noWrap,omitempty"`
	TcMar      *TcMar         `xml:"tcMar,omitempty"`
	TextDirection *TextDirection `xml:"textDirection,omitempty"`
}

// GridSpan represents cell column span.
type GridSpan struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// VMerge represents vertical merge.
type VMerge struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
}

// TcBorders represents cell borders.
type TcBorders struct {
	Top    *Border `xml:"top,omitempty"`
	Left   *Border `xml:"left,omitempty"`
	Bottom *Border `xml:"bottom,omitempty"`
	Right  *Border `xml:"right,omitempty"`
}

// Shd represents shading.
type Shd struct {
	Val   string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	Color string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main color,attr,omitempty"`
	Fill  string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main fill,attr,omitempty"`
}

// VAlign represents vertical alignment.
type VAlign struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// TextDirection represents cell text direction.
type TextDirection struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// TcMar represents cell margins.
type TcMar struct {
	Top    *TblWidth `xml:"top,omitempty"`
	Left   *TblWidth `xml:"left,omitempty"`
	Bottom *TblWidth `xml:"bottom,omitempty"`
	Right  *TblWidth `xml:"right,omitempty"`
}

// UnmarshalXML implements custom unmarshaling for table cells to properly parse nested elements.
func (tc *Tc) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	tc.XMLName = start.Name
	tc.Content = nil

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcPr":
				var tcPr TcPr
				if err := d.DecodeElement(&tcPr, &t); err != nil {
					return err
				}
				tc.TcPr = &tcPr
			case "p":
				var p P
				if err := d.DecodeElement(&p, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, &p)
			case "tbl":
				var tbl Tbl
				if err := d.DecodeElement(&tbl, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, &tbl)
			default:
				// Skip unknown elements
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t == start.End() {
				return nil
			}
		}
	}
}

// MarshalXML implements custom marshaling for table cells.
func (tc *Tc) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "tc"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if tc.TcPr != nil {
		if err := e.EncodeElement(tc.TcPr, xml.StartElement{Name: xml.Name{Space: NS, Local: "tcPr"}}); err != nil {
			return err
		}
	}

	for _, elem := range tc.Content {
		switch v := elem.(type) {
		case *P:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "p"}}); err != nil {
				return err
			}
		case *Tbl:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "tbl"}}); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(start.End())
}
