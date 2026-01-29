package dml

import "encoding/xml"

// Color-related types

// NoFill represents no fill.
type NoFill struct{}

// SolidFill represents solid fill.
type SolidFill struct {
	SrgbClr   *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
	PrstClr   *PrstClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main prstClr,omitempty"`
	SysClr    *SysClr    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main sysClr,omitempty"`
}

// SrgbClr represents an sRGB color.
type SrgbClr struct {
	Val    string       `xml:"val,attr"` // Hex RGB (e.g., "FF0000")
	Lummod *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lumMod,omitempty"`
	Lumoff *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lumOff,omitempty"`
	Tint   *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tint,omitempty"`
	Shade  *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main shade,omitempty"`
	SatMod *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main satMod,omitempty"`
	Alpha  *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main alpha,omitempty"`
}

// SchemeClr represents a theme scheme color.
type SchemeClr struct {
	Val    string       `xml:"val,attr"` // dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
	Lummod *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lumMod,omitempty"`
	Lumoff *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lumOff,omitempty"`
	Tint   *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tint,omitempty"`
	Shade  *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main shade,omitempty"`
	SatMod *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main satMod,omitempty"`
	Alpha  *ColorMod    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main alpha,omitempty"`
}

// PrstClr represents a preset color.
type PrstClr struct {
	Val string `xml:"val,attr"` // black, white, red, etc.
}

// SysClr represents a system color.
type SysClr struct {
	Val     string `xml:"val,attr"`               // windowText, window, etc.
	LastClr string `xml:"lastClr,attr,omitempty"` // Last rendered color
}

// ColorMod represents a color modification.
type ColorMod struct {
	Val int `xml:"val,attr"` // Percentage * 1000
}

// GradFill represents gradient fill.
type GradFill struct {
	Flip     string   `xml:"flip,attr,omitempty"` // none, x, y, xy
	RotWithShape *bool `xml:"rotWithShape,attr,omitempty"`
	GsLst    *GsLst   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gsLst,omitempty"`
	Lin      *Lin     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lin,omitempty"`
	Path     *GradPath `xml:"http://schemas.openxmlformats.org/drawingml/2006/main path,omitempty"`
}

// GsLst is a gradient stop list.
type GsLst struct {
	Gs []*Gs `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gs,omitempty"`
}

// Gs represents a gradient stop.
type Gs struct {
	Pos       int        `xml:"pos,attr"` // Position 0-100000
	SrgbClr   *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

// Lin represents linear gradient properties.
type Lin struct {
	Ang    int64 `xml:"ang,attr"`              // Angle in 60000ths of a degree
	Scaled *bool `xml:"scaled,attr,omitempty"`
}

// GradPath represents gradient path properties.
type GradPath struct {
	Path string `xml:"path,attr"` // circle, rect, shape
}

// BlipFill represents picture fill.
type BlipFill struct {
	Dpi        int        `xml:"dpi,attr,omitempty"`
	RotWithShape *bool    `xml:"rotWithShape,attr,omitempty"`
	Blip       *Blip      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blip,omitempty"`
	SrcRect    *SrcRect   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srcRect,omitempty"`
	Tile       *Tile      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tile,omitempty"`
	Stretch    *Stretch   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main stretch,omitempty"`
}

// Blip represents an image reference.
type Blip struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blip"`
	Embed   string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr,omitempty"`
	Link    string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships link,attr,omitempty"`
	CState  string   `xml:"cstate,attr,omitempty"` // none, print, screen, email
}

// SrcRect represents source rectangle for cropping.
type SrcRect struct {
	L int `xml:"l,attr,omitempty"` // Left (percentage * 1000)
	T int `xml:"t,attr,omitempty"` // Top
	R int `xml:"r,attr,omitempty"` // Right
	B int `xml:"b,attr,omitempty"` // Bottom
}

// Tile represents tile settings for fill.
type Tile struct {
	Tx    int64  `xml:"tx,attr,omitempty"`    // Horizontal offset
	Ty    int64  `xml:"ty,attr,omitempty"`    // Vertical offset
	Sx    int    `xml:"sx,attr,omitempty"`    // Horizontal scale (percentage * 1000)
	Sy    int    `xml:"sy,attr,omitempty"`    // Vertical scale
	Flip  string `xml:"flip,attr,omitempty"`  // none, x, y, xy
	Algn  string `xml:"algn,attr,omitempty"`  // Alignment
}

// Stretch represents stretch fill settings.
type Stretch struct {
	FillRect *FillRect `xml:"http://schemas.openxmlformats.org/drawingml/2006/main fillRect,omitempty"`
}

// FillRect represents fill rectangle.
type FillRect struct {
	L int `xml:"l,attr,omitempty"`
	T int `xml:"t,attr,omitempty"`
	R int `xml:"r,attr,omitempty"`
	B int `xml:"b,attr,omitempty"`
}

// Ln represents line properties.
type Ln struct {
	W          int64      `xml:"w,attr,omitempty"`          // Width in EMUs
	Cap        string     `xml:"cap,attr,omitempty"`        // flat, sq, rnd
	Cmpd       string     `xml:"cmpd,attr,omitempty"`       // sng, dbl, thickThin, etc.
	Algn       string     `xml:"algn,attr,omitempty"`       // ctr, in
	NoFill     *NoFill    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main noFill,omitempty"`
	SolidFill  *SolidFill `xml:"http://schemas.openxmlformats.org/drawingml/2006/main solidFill,omitempty"`
	GradFill   *GradFill  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gradFill,omitempty"`
	PrstDash   *PrstDash  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main prstDash,omitempty"`
	Round      *Round     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main round,omitempty"`
	Bevel      *Bevel     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main bevel,omitempty"`
	Miter      *Miter     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main miter,omitempty"`
	HeadEnd    *LineEnd   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main headEnd,omitempty"`
	TailEnd    *LineEnd   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main tailEnd,omitempty"`
}

// PrstDash represents preset dash.
type PrstDash struct {
	Val string `xml:"val,attr"` // solid, dot, dash, lgDash, etc.
}

// Round represents round line join.
type Round struct{}

// Bevel represents bevel line join.
type Bevel struct{}

// Miter represents miter line join.
type Miter struct {
	Lim int `xml:"lim,attr,omitempty"` // Miter limit (percentage * 1000)
}

// LineEnd represents line end (arrow).
type LineEnd struct {
	Type string `xml:"type,attr,omitempty"` // none, triangle, stealth, diamond, oval, arrow
	W    string `xml:"w,attr,omitempty"`    // sm, med, lg
	Len  string `xml:"len,attr,omitempty"`  // sm, med, lg
}

// EffectLst represents effect list.
type EffectLst struct {
	OuterShdw *OuterShdw `xml:"http://schemas.openxmlformats.org/drawingml/2006/main outerShdw,omitempty"`
	InnerShdw *InnerShdw `xml:"http://schemas.openxmlformats.org/drawingml/2006/main innerShdw,omitempty"`
	Reflection *Reflection `xml:"http://schemas.openxmlformats.org/drawingml/2006/main reflection,omitempty"`
	Glow      *Glow      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main glow,omitempty"`
	SoftEdge  *SoftEdge  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main softEdge,omitempty"`
}

// OuterShdw represents outer shadow.
type OuterShdw struct {
	BlurRad  int64      `xml:"blurRad,attr,omitempty"`
	Dist     int64      `xml:"dist,attr,omitempty"`
	Dir      int64      `xml:"dir,attr,omitempty"` // Angle in 60000ths of a degree
	Sx       int        `xml:"sx,attr,omitempty"`
	Sy       int        `xml:"sy,attr,omitempty"`
	Kx       int64      `xml:"kx,attr,omitempty"`
	Ky       int64      `xml:"ky,attr,omitempty"`
	Algn     string     `xml:"algn,attr,omitempty"`
	RotWithShape *bool  `xml:"rotWithShape,attr,omitempty"`
	SrgbClr  *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

// InnerShdw represents inner shadow.
type InnerShdw struct {
	BlurRad  int64      `xml:"blurRad,attr,omitempty"`
	Dist     int64      `xml:"dist,attr,omitempty"`
	Dir      int64      `xml:"dir,attr,omitempty"`
	SrgbClr  *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

// Reflection represents reflection effect.
type Reflection struct {
	BlurRad  int64  `xml:"blurRad,attr,omitempty"`
	StA      int    `xml:"stA,attr,omitempty"` // Start alpha
	EndA     int    `xml:"endA,attr,omitempty"` // End alpha
	Dist     int64  `xml:"dist,attr,omitempty"`
	Dir      int64  `xml:"dir,attr,omitempty"`
	FadeDir  int64  `xml:"fadeDir,attr,omitempty"`
	Sx       int    `xml:"sx,attr,omitempty"`
	Sy       int    `xml:"sy,attr,omitempty"`
	Kx       int64  `xml:"kx,attr,omitempty"`
	Ky       int64  `xml:"ky,attr,omitempty"`
	Algn     string `xml:"algn,attr,omitempty"`
	RotWithShape *bool `xml:"rotWithShape,attr,omitempty"`
}

// Glow represents glow effect.
type Glow struct {
	Rad      int64      `xml:"rad,attr"`
	SrgbClr  *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

// SoftEdge represents soft edge effect.
type SoftEdge struct {
	Rad int64 `xml:"rad,attr"`
}

// Scheme color values.
const (
	SchemeClrDk1      = "dk1"
	SchemeClrLt1      = "lt1"
	SchemeClrDk2      = "dk2"
	SchemeClrLt2      = "lt2"
	SchemeClrAccent1  = "accent1"
	SchemeClrAccent2  = "accent2"
	SchemeClrAccent3  = "accent3"
	SchemeClrAccent4  = "accent4"
	SchemeClrAccent5  = "accent5"
	SchemeClrAccent6  = "accent6"
	SchemeClrHlink    = "hlink"
	SchemeClrFolHlink = "folHlink"
)
