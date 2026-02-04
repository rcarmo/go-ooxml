package sml

import "encoding/xml"

// StyleSheet represents the styles part.
type StyleSheet struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	NumFmts       *NumFmts       `xml:"numFmts,omitempty"`
	Fonts         *Fonts         `xml:"fonts,omitempty"`
	Fills         *Fills         `xml:"fills,omitempty"`
	Borders       *Borders       `xml:"borders,omitempty"`
	CellStyleXfs  *CellStyleXfs  `xml:"cellStyleXfs,omitempty"`
	CellXfs       *CellXfs       `xml:"cellXfs,omitempty"`
	CellStyles    *CellStyles    `xml:"cellStyles,omitempty"`
	TableStyles   *TableStyles   `xml:"tableStyles,omitempty"`
	Colors        *Colors        `xml:"colors,omitempty"`
}

// NumFmts represents number formats.
type NumFmts struct {
	Count  int       `xml:"count,attr,omitempty"`
	NumFmt []*NumFmt `xml:"numFmt,omitempty"`
}

// NumFmt represents a number format.
type NumFmt struct {
	NumFmtID   int    `xml:"numFmtId,attr"`
	FormatCode string `xml:"formatCode,attr"`
}

// Fonts represents fonts collection.
type Fonts struct {
	Count int     `xml:"count,attr,omitempty"`
	Font  []*Font `xml:"font,omitempty"`
}

// Font represents font settings.
type Font struct {
	B      *BoolVal    `xml:"b,omitempty"`
	I      *BoolVal    `xml:"i,omitempty"`
	Sz     *FontSize   `xml:"sz,omitempty"`
	Color  *Color      `xml:"color,omitempty"`
	Name   *FontName   `xml:"name,omitempty"`
	Family *FontFamily `xml:"family,omitempty"`
	Scheme *FontScheme `xml:"scheme,omitempty"`
}

// BoolVal represents a boolean value attribute.
type BoolVal struct {
	Val string `xml:"val,attr,omitempty"`
}

// FontSize represents font size.
type FontSize struct {
	Val float64 `xml:"val,attr"`
}

// FontName represents font name.
type FontName struct {
	Val string `xml:"val,attr"`
}

// FontFamily represents font family.
type FontFamily struct {
	Val int `xml:"val,attr,omitempty"`
}

// FontScheme represents font scheme.
type FontScheme struct {
	Val string `xml:"val,attr,omitempty"`
}

// Color represents a color definition.
type Color struct {
	RGB     string   `xml:"rgb,attr,omitempty"`
	Theme   *int     `xml:"theme,attr,omitempty"`
	Tint    *float64 `xml:"tint,attr,omitempty"`
	Indexed *int     `xml:"indexed,attr,omitempty"`
	Auto    *bool    `xml:"auto,attr,omitempty"`
}

// Fills represents fills collection.
type Fills struct {
	Count int     `xml:"count,attr,omitempty"`
	Fill  []*Fill `xml:"fill,omitempty"`
}

// Fill represents a fill.
type Fill struct {
	PatternFill *PatternFill `xml:"patternFill,omitempty"`
}

// PatternFill represents a pattern fill.
type PatternFill struct {
	PatternType string `xml:"patternType,attr,omitempty"`
	FgColor     *Color `xml:"fgColor,omitempty"`
	BgColor     *Color `xml:"bgColor,omitempty"`
}

// Borders represents borders collection.
type Borders struct {
	Count  int       `xml:"count,attr,omitempty"`
	Border []*Border `xml:"border,omitempty"`
}

// Border represents a border.
type Border struct {
	Left     *BorderSide `xml:"left,omitempty"`
	Right    *BorderSide `xml:"right,omitempty"`
	Top      *BorderSide `xml:"top,omitempty"`
	Bottom   *BorderSide `xml:"bottom,omitempty"`
	Diagonal *BorderSide `xml:"diagonal,omitempty"`
}

// BorderSide represents a border side.
type BorderSide struct {
	Style string `xml:"style,attr,omitempty"`
	Color *Color `xml:"color,omitempty"`
}

// CellStyleXfs represents cell style XFs.
type CellStyleXfs struct {
	Count int   `xml:"count,attr,omitempty"`
	Xf    []*Xf `xml:"xf,omitempty"`
}

// CellXfs represents cell XFs.
type CellXfs struct {
	Count int   `xml:"count,attr,omitempty"`
	Xf    []*Xf `xml:"xf,omitempty"`
}

// Xf represents a format record.
type Xf struct {
	NumFmtID          int        `xml:"numFmtId,attr,omitempty"`
	FontID            int        `xml:"fontId,attr,omitempty"`
	FillID            int        `xml:"fillId,attr,omitempty"`
	BorderID          int        `xml:"borderId,attr,omitempty"`
	XFID              int        `xml:"xfId,attr,omitempty"`
	ApplyAlignment    *bool      `xml:"applyAlignment,attr,omitempty"`
	ApplyNumberFormat *bool      `xml:"applyNumberFormat,attr,omitempty"`
	Alignment         *Alignment `xml:"alignment,omitempty"`
}

// Alignment represents cell alignment.
type Alignment struct {
	Horizontal string `xml:"horizontal,attr,omitempty"`
	Vertical   string `xml:"vertical,attr,omitempty"`
}

// CellStyles represents named styles.
type CellStyles struct {
	Count     int         `xml:"count,attr,omitempty"`
	CellStyle []*CellStyle `xml:"cellStyle,omitempty"`
}

// CellStyle represents a named style.
type CellStyle struct {
	Name      string `xml:"name,attr,omitempty"`
	XFID      int    `xml:"xfId,attr,omitempty"`
	BuiltinID int    `xml:"builtinId,attr,omitempty"`
	Hidden    *bool  `xml:"hidden,attr,omitempty"`
}

// TableStyles represents table styles.
type TableStyles struct {
	Count              int    `xml:"count,attr,omitempty"`
	DefaultTableStyle  string `xml:"defaultTableStyle,attr,omitempty"`
	DefaultPivotStyle  string `xml:"defaultPivotStyle,attr,omitempty"`
}

// Colors represents colors collection.
type Colors struct {
	IndexedColors *IndexedColors `xml:"indexedColors,omitempty"`
}

// IndexedColors represents indexed colors.
type IndexedColors struct {
	RGBColor []*RGBColor `xml:"rgbColor,omitempty"`
}

// RGBColor represents an indexed color value.
type RGBColor struct {
	RGB string `xml:"rgb,attr,omitempty"`
}
