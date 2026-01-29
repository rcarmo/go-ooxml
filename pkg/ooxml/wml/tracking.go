package wml

import "encoding/xml"

// Ins represents an insertion (tracked change).
type Ins struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main ins"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	Content []interface{} `xml:",any"`
}

// Del represents a deletion (tracked change).
type Del struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main del"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	Content []interface{} `xml:",any"`
}

// DelText represents deleted text.
type DelText struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main delText"`
	Space   string   `xml:"http://www.w3.org/XML/1998/namespace space,attr,omitempty"`
	Text    string   `xml:",chardata"`
}

// MoveTo represents moved-to content.
type MoveTo struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main moveTo"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	Content []interface{} `xml:",any"`
}

// MoveFrom represents moved-from content.
type MoveFrom struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main moveFrom"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	Content []interface{} `xml:",any"`
}

// MoveToRangeStart marks the start of moved-to range.
type MoveToRangeStart struct {
	XMLName    xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main moveToRangeStart"`
	ID         int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Name       string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main name,attr,omitempty"`
	Author     string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date       string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	ColFirst   *int     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main colFirst,attr,omitempty"`
	ColLast    *int     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main colLast,attr,omitempty"`
	DisplacedByCustomXml string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main displacedByCustomXml,attr,omitempty"`
}

// MoveToRangeEnd marks the end of moved-to range.
type MoveToRangeEnd struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main moveToRangeEnd"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
}

// MoveFromRangeStart marks the start of moved-from range.
type MoveFromRangeStart struct {
	XMLName    xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main moveFromRangeStart"`
	ID         int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Name       string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main name,attr,omitempty"`
	Author     string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date       string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	ColFirst   *int     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main colFirst,attr,omitempty"`
	ColLast    *int     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main colLast,attr,omitempty"`
	DisplacedByCustomXml string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main displacedByCustomXml,attr,omitempty"`
}

// MoveFromRangeEnd marks the end of moved-from range.
type MoveFromRangeEnd struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main moveFromRangeEnd"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
}

// RPrChange represents run properties change.
type RPrChange struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main rPrChange"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	RPr     *RPr     `xml:"rPr,omitempty"`
}

// PPrChange represents paragraph properties change.
type PPrChange struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main pPrChange"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	PPr     *PPr     `xml:"pPr,omitempty"`
}

// SectPrChange represents section properties change.
type SectPrChange struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sectPrChange"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	SectPr  *SectPr  `xml:"sectPr,omitempty"`
}

// TblPrChange represents table properties change.
type TblPrChange struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tblPrChange"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	TblPr   *TblPr   `xml:"tblPr,omitempty"`
}

// TrPrChange represents table row properties change.
type TrPrChange struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main trPrChange"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	TrPr    *TrPr    `xml:"trPr,omitempty"`
}

// TcPrChange represents table cell properties change.
type TcPrChange struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tcPrChange"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Author  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main author,attr,omitempty"`
	Date    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main date,attr,omitempty"`
	TcPr    *TcPr    `xml:"tcPr,omitempty"`
}
