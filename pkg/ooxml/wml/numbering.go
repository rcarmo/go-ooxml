package wml

import "encoding/xml"

// Numbering represents the numbering definitions part.
type Numbering struct {
	XMLName     xml.Name     `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main numbering"`
	AbstractNum []*AbstractNum `xml:"abstractNum,omitempty"`
	Num         []*Num        `xml:"num,omitempty"`
}

// AbstractNum represents an abstract numbering definition.
type AbstractNum struct {
	AbstractNumID int    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main abstractNumId,attr"`
	Nsid          *Nsid  `xml:"nsid,omitempty"`
	MultiLevelType *MultiLevelType `xml:"multiLevelType,omitempty"`
	Tmpl          *Tmpl  `xml:"tmpl,omitempty"`
	Lvl           []*Lvl `xml:"lvl,omitempty"`
}

// Nsid represents a numbering definition ID.
type Nsid struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// MultiLevelType represents the multi-level type.
type MultiLevelType struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Tmpl represents a template ID.
type Tmpl struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Lvl represents a numbering level definition.
type Lvl struct {
	Ilvl       int         `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main ilvl,attr"`
	Tplc       string      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tplc,attr,omitempty"`
	Start      *NumStart   `xml:"start,omitempty"`
	NumFmt     *NumFmt     `xml:"numFmt,omitempty"`
	LvlRestart *LvlRestart `xml:"lvlRestart,omitempty"`
	Suff       *Suff       `xml:"suff,omitempty"`
	LvlText    *LvlText    `xml:"lvlText,omitempty"`
	LvlJc      *LvlJc      `xml:"lvlJc,omitempty"`
	PPr        *PPr        `xml:"pPr,omitempty"`
	RPr        *RPr        `xml:"rPr,omitempty"`
}

// NumStart represents the starting value.
type NumStart struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// NumFmt represents the number format.
type NumFmt struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// LvlRestart represents level restart behavior.
type LvlRestart struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Suff represents the suffix after the number.
type Suff struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// LvlText represents the level text.
type LvlText struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// LvlJc represents level justification.
type LvlJc struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Num represents a numbering definition instance.
type Num struct {
	NumID        int           `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main numId,attr"`
	AbstractNumID *AbstractNumIDRef `xml:"abstractNumId,omitempty"`
	LvlOverride  []*LvlOverride `xml:"lvlOverride,omitempty"`
}

// AbstractNumIDRef references an abstract numbering definition.
type AbstractNumIDRef struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// LvlOverride represents a level override.
type LvlOverride struct {
	Ilvl      int       `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main ilvl,attr"`
	StartOverride *StartOverride `xml:"startOverride,omitempty"`
	Lvl       *Lvl      `xml:"lvl,omitempty"`
}

// StartOverride represents a start value override.
type StartOverride struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Common number formats.
const (
	NumFmtDecimal        = "decimal"
	NumFmtUpperRoman     = "upperRoman"
	NumFmtLowerRoman     = "lowerRoman"
	NumFmtUpperLetter    = "upperLetter"
	NumFmtLowerLetter    = "lowerLetter"
	NumFmtBullet         = "bullet"
	NumFmtNone           = "none"
)
