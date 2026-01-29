// Package wml provides WordprocessingML types for OOXML documents.
package wml

import "encoding/xml"

// Namespaces used in WordprocessingML documents.
const (
	NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	NSR  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	NSW  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	NSW14 = "http://schemas.microsoft.com/office/word/2010/wordml"
)

// Document is the root element of a WordprocessingML document.
type Document struct {
	XMLName             xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main document"`
	MCIgnorable         string   `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 Ignorable,attr,omitempty"`
	Body                *Body    `xml:"body"`
}

// Body contains the block-level content of the document.
type Body struct {
	Content []interface{} `xml:"-"` // Handled by custom unmarshaler
	SectPr  *SectPr       `xml:"sectPr,omitempty"`
}

// UnmarshalXML implements custom XML unmarshaling for Body.
func (b *Body) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "p":
				p := &P{}
				if err := d.DecodeElement(p, &t); err != nil {
					return err
				}
				b.Content = append(b.Content, p)
			case "tbl":
				tbl := &Tbl{}
				if err := d.DecodeElement(tbl, &t); err != nil {
					return err
				}
				b.Content = append(b.Content, tbl)
			case "sectPr":
				b.SectPr = &SectPr{}
				if err := d.DecodeElement(b.SectPr, &t); err != nil {
					return err
				}
			default:
				// Skip unknown elements
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name == start.Name {
				return nil
			}
		}
	}
}

// MarshalXML implements custom XML marshaling for Body.
func (b *Body) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:body"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, elem := range b.Content {
		if err := e.Encode(elem); err != nil {
			return err
		}
	}

	if b.SectPr != nil {
		if err := e.Encode(b.SectPr); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// SectPr represents section properties.
type SectPr struct {
	XMLName     xml.Name    `xml:"sectPr"`
	HeaderRefs  []HeaderRef `xml:"headerReference,omitempty"`
	FooterRefs  []FooterRef `xml:"footerReference,omitempty"`
	PgSz        *PgSz       `xml:"pgSz,omitempty"`
	PgMar       *PgMar      `xml:"pgMar,omitempty"`
	Cols        *Cols       `xml:"cols,omitempty"`
	DocGrid     *DocGrid    `xml:"docGrid,omitempty"`
}

// HeaderRef references a header part.
type HeaderRef struct {
	XMLName xml.Name `xml:"headerReference"`
	Type    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main type,attr"`
	ID      string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// FooterRef references a footer part.
type FooterRef struct {
	XMLName xml.Name `xml:"footerReference"`
	Type    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main type,attr"`
	ID      string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// PgSz represents page size.
type PgSz struct {
	W      int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main w,attr,omitempty"`
	H      int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main h,attr,omitempty"`
	Orient string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main orient,attr,omitempty"`
}

// PgMar represents page margins.
type PgMar struct {
	Top    int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main top,attr,omitempty"`
	Right  int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main right,attr,omitempty"`
	Bottom int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main bottom,attr,omitempty"`
	Left   int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main left,attr,omitempty"`
	Header int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main header,attr,omitempty"`
	Footer int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main footer,attr,omitempty"`
	Gutter int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main gutter,attr,omitempty"`
}

// Cols represents column settings.
type Cols struct {
	Space int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main space,attr,omitempty"`
}

// DocGrid represents document grid settings.
type DocGrid struct {
	LinePitch int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main linePitch,attr,omitempty"`
}
