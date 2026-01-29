package sml

import "encoding/xml"

// Cell represents a cell.
type Cell struct {
	XMLName xml.Name `xml:"c"`
	R       string   `xml:"r,attr,omitempty"` // Reference (e.g., "A1")
	S       int      `xml:"s,attr,omitempty"` // Style index
	T       string   `xml:"t,attr,omitempty"` // Type: s=shared string, n=number, b=boolean, e=error, str=string
	F       *Formula `xml:"f,omitempty"`      // Formula
	V       string   `xml:"v,omitempty"`      // Value
	Is      *Is      `xml:"is,omitempty"`     // Inline string
}

// Formula represents a cell formula.
type Formula struct {
	T        string `xml:"t,attr,omitempty"`   // Type: normal, array, dataTable, shared
	Ref      string `xml:"ref,attr,omitempty"` // Range for array/shared formulas
	SI       *int   `xml:"si,attr,omitempty"`  // Shared formula index
	Content  string `xml:",chardata"`
}

// Is represents an inline string.
type Is struct {
	T string `xml:"t,omitempty"`
	R []*RT  `xml:"r,omitempty"` // Rich text runs
}

// RT represents a rich text run within an inline string.
type RT struct {
	RPr *RTRPr `xml:"rPr,omitempty"`
	T   string `xml:"t"`
}

// RTRPr represents rich text run properties.
type RTRPr struct {
	B       *RTBool  `xml:"b,omitempty"`
	I       *RTBool  `xml:"i,omitempty"`
	Strike  *RTBool  `xml:"strike,omitempty"`
	Condense *RTBool `xml:"condense,omitempty"`
	Extend  *RTBool  `xml:"extend,omitempty"`
	Outline *RTBool  `xml:"outline,omitempty"`
	Shadow  *RTBool  `xml:"shadow,omitempty"`
	U       *RTU     `xml:"u,omitempty"`
	VertAlign *RTVertAlign `xml:"vertAlign,omitempty"`
	Sz      *RTSz    `xml:"sz,omitempty"`
	Color   *RTColor `xml:"color,omitempty"`
	RFont   *RTRFont `xml:"rFont,omitempty"`
	Family  *RTFamily `xml:"family,omitempty"`
	Charset *RTCharset `xml:"charset,omitempty"`
	Scheme  *RTScheme `xml:"scheme,omitempty"`
}

// RTBool represents a boolean property.
type RTBool struct {
	Val *bool `xml:"val,attr,omitempty"`
}

// RTU represents underline.
type RTU struct {
	Val string `xml:"val,attr,omitempty"`
}

// RTVertAlign represents vertical alignment.
type RTVertAlign struct {
	Val string `xml:"val,attr"` // superscript, subscript, baseline
}

// RTSz represents font size.
type RTSz struct {
	Val float64 `xml:"val,attr"`
}

// RTColor represents color.
type RTColor struct {
	Auto    *bool   `xml:"auto,attr,omitempty"`
	Indexed int     `xml:"indexed,attr,omitempty"`
	RGB     string  `xml:"rgb,attr,omitempty"`
	Theme   int     `xml:"theme,attr,omitempty"`
	Tint    float64 `xml:"tint,attr,omitempty"`
}

// RTRFont represents font name.
type RTRFont struct {
	Val string `xml:"val,attr"`
}

// RTFamily represents font family.
type RTFamily struct {
	Val int `xml:"val,attr"`
}

// RTCharset represents character set.
type RTCharset struct {
	Val int `xml:"val,attr"`
}

// RTScheme represents font scheme.
type RTScheme struct {
	Val string `xml:"val,attr"` // major, minor, none
}

// Cell type constants.
const (
	CellTypeBool         = "b"
	CellTypeNumber       = "n"
	CellTypeError        = "e"
	CellTypeSharedString = "s"
	CellTypeString       = "str"
	CellTypeInlineString = "inlineStr"
)
