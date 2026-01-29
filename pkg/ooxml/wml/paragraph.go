package wml

import "encoding/xml"

// P represents a paragraph.
type P struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	PPr     *PPr     `xml:"pPr,omitempty"`
	Content []interface{} `xml:",any"`
}

// PPr represents paragraph properties.
type PPr struct {
	XMLName    xml.Name    `xml:"pPr"`
	PStyle     *PStyle     `xml:"pStyle,omitempty"`
	KeepNext   *OnOff      `xml:"keepNext,omitempty"`
	KeepLines  *OnOff      `xml:"keepLines,omitempty"`
	PageBreakBefore *OnOff `xml:"pageBreakBefore,omitempty"`
	Spacing    *Spacing    `xml:"spacing,omitempty"`
	Ind        *Ind        `xml:"ind,omitempty"`
	Jc         *Jc         `xml:"jc,omitempty"`
	RPr        *RPr        `xml:"rPr,omitempty"`
	NumPr      *NumPr      `xml:"numPr,omitempty"`
	OutlineLvl *OutlineLvl `xml:"outlineLvl,omitempty"`
}

// PStyle references a paragraph style.
type PStyle struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// OnOff represents a boolean toggle.
type OnOff struct {
	Val *bool `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
}

// Spacing represents paragraph spacing.
type Spacing struct {
	Before   *int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main before,attr,omitempty"`
	After    *int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main after,attr,omitempty"`
	Line     *int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main line,attr,omitempty"`
	LineRule *string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main lineRule,attr,omitempty"`
}

// Ind represents paragraph indentation.
type Ind struct {
	Left      *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main left,attr,omitempty"`
	Right     *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main right,attr,omitempty"`
	FirstLine *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main firstLine,attr,omitempty"`
	Hanging   *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hanging,attr,omitempty"`
}

// Jc represents paragraph justification.
type Jc struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// NumPr represents numbering properties.
type NumPr struct {
	Ilvl  *Ilvl  `xml:"ilvl,omitempty"`
	NumID *NumID `xml:"numId,omitempty"`
}

// Ilvl represents indentation level for numbering.
type Ilvl struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// NumID references a numbering definition.
type NumID struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// OutlineLvl represents outline level.
type OutlineLvl struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Enabled returns true if the OnOff element is set and enabled.
func (o *OnOff) Enabled() bool {
	if o == nil {
		return false
	}
	if o.Val == nil {
		return true // presence without val means true
	}
	return *o.Val
}

// NewOnOff creates an OnOff element with the given value.
func NewOnOff(val bool) *OnOff {
	return &OnOff{Val: &val}
}

// NewOnOffEnabled creates an enabled OnOff element (no val attribute).
func NewOnOffEnabled() *OnOff {
	return &OnOff{}
}
