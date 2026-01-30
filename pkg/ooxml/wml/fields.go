package wml

import "encoding/xml"

// Field character types.
const (
	FldCharBegin    = "begin"
	FldCharSeparate = "separate"
	FldCharEnd      = "end"
)

// FldChar represents a field character run element.
type FldChar struct {
	XMLName      xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main fldChar"`
	FldCharType  string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main fldCharType,attr"`
	Dirty        *bool    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main dirty,attr,omitempty"`
	Lock         *bool    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main lock,attr,omitempty"`
}

// InstrText represents field instruction text.
type InstrText struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main instrText"`
	Space   string   `xml:"http://www.w3.org/XML/1998/namespace space,attr,omitempty"`
	Text    string   `xml:",chardata"`
}

// NewInstrText creates a new instruction text element.
func NewInstrText(text string) *InstrText {
	it := &InstrText{Text: text}
	if len(text) > 0 && (text[0] == ' ' || text[len(text)-1] == ' ' || containsMultipleSpaces(text)) {
		it.Space = "preserve"
	}
	return it
}
