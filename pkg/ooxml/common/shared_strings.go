// Package common provides shared types used across OOXML formats.
package common

import "encoding/xml"

// SharedStrings represents the shared strings part (used by Excel).
type SharedStrings struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr,omitempty"`
	UniqueCount int      `xml:"uniqueCount,attr,omitempty"`
	SI          []*SI    `xml:"si,omitempty"`
}

// SI represents a shared string item.
type SI struct {
	T string `xml:"t,omitempty"`
	R []*RT  `xml:"r,omitempty"` // Rich text
}

// RT represents rich text.
type RT struct {
	RPr *RTRPr `xml:"rPr,omitempty"`
	T   string `xml:"t"`
}

// RTRPr represents rich text run properties.
type RTRPr struct {
	B      *RTBool  `xml:"b,omitempty"`
	I      *RTBool  `xml:"i,omitempty"`
	Strike *RTBool  `xml:"strike,omitempty"`
	U      *RTU     `xml:"u,omitempty"`
	Sz     *RTSz    `xml:"sz,omitempty"`
	Color  *RTColor `xml:"color,omitempty"`
	RFont  *RTRFont `xml:"rFont,omitempty"`
	Family *RTFamily `xml:"family,omitempty"`
	Scheme *RTScheme `xml:"scheme,omitempty"`
}

// RTBool represents a boolean value.
type RTBool struct {
	Val *bool `xml:"val,attr,omitempty"`
}

// RTU represents underline.
type RTU struct {
	Val string `xml:"val,attr,omitempty"`
}

// RTSz represents font size.
type RTSz struct {
	Val float64 `xml:"val,attr"`
}

// RTColor represents text color.
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

// RTScheme represents font scheme.
type RTScheme struct {
	Val string `xml:"val,attr"` // major, minor, none
}

// GetString returns the string at the given index.
func (ss *SharedStrings) GetString(index int) string {
	if index < 0 || index >= len(ss.SI) {
		return ""
	}
	si := ss.SI[index]
	if si.T != "" {
		return si.T
	}
	// Handle rich text
	var result string
	for _, r := range si.R {
		result += r.T
	}
	return result
}

// AddString adds a string and returns its index.
func (ss *SharedStrings) AddString(s string) int {
	// Check if already exists
	for i, si := range ss.SI {
		if si.T == s {
			return i
		}
	}
	// Add new
	ss.SI = append(ss.SI, &SI{T: s})
	ss.Count++
	ss.UniqueCount = len(ss.SI)
	return len(ss.SI) - 1
}
