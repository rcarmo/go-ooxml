// Package dml provides DrawingML types shared across OOXML formats.
package dml

import "encoding/xml"

// Namespaces used in DrawingML.
const (
	NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
)

// Sp represents a shape.
type Sp struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main sp"`
	NvSpPr  *NvSpPr  `xml:"nvSpPr"`
	SpPr    *SpPr    `xml:"spPr"`
	Style   *ShapeStyle `xml:"style,omitempty"`
	TxBody  *TxBody  `xml:"txBody,omitempty"`
}

// NvSpPr represents non-visual shape properties.
type NvSpPr struct {
	CNvPr   *CNvPr   `xml:"cNvPr"`
	CNvSpPr *CNvSpPr `xml:"cNvSpPr"`
	NvPr    *NvPr    `xml:"nvPr,omitempty"`
}

// CNvPr represents common non-visual properties.
type CNvPr struct {
	ID     int     `xml:"id,attr"`
	Name   string  `xml:"name,attr"`
	Descr  string  `xml:"descr,attr,omitempty"`
	Title  string  `xml:"title,attr,omitempty"`
	Hidden *bool   `xml:"hidden,attr,omitempty"`
	HlinkClick *HlinkClick `xml:"hlinkClick,omitempty"`
}

// HlinkClick represents a hyperlink.
type HlinkClick struct {
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	Action  string `xml:"action,attr,omitempty"`
	Tooltip string `xml:"tooltip,attr,omitempty"`
}

// CNvSpPr represents non-visual shape drawing properties.
type CNvSpPr struct {
	TxBox *bool `xml:"txBox,attr,omitempty"` // Is text box
}

// NvPr represents non-visual properties.
type NvPr struct {
	Ph *Ph `xml:"ph,omitempty"`
}

// Ph represents placeholder info.
type Ph struct {
	Type string `xml:"type,attr,omitempty"`
	Idx  *int   `xml:"idx,attr,omitempty"`
}

// SpPr represents shape properties.
type SpPr struct {
	Xfrm       *Xfrm      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main xfrm,omitempty"`
	PrstGeom   *PrstGeom  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main prstGeom,omitempty"`
	CustGeom   *CustGeom  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main custGeom,omitempty"`
	NoFill     *NoFill    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main noFill,omitempty"`
	SolidFill  *SolidFill `xml:"http://schemas.openxmlformats.org/drawingml/2006/main solidFill,omitempty"`
	GradFill   *GradFill  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gradFill,omitempty"`
	BlipFill   *BlipFill  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blipFill,omitempty"`
	Ln         *Ln        `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ln,omitempty"`
	EffectLst  *EffectLst `xml:"http://schemas.openxmlformats.org/drawingml/2006/main effectLst,omitempty"`
}

// Xfrm represents 2D transform.
type Xfrm struct {
	Rot   int64 `xml:"rot,attr,omitempty"`   // Rotation in 60000ths of a degree
	FlipH *bool `xml:"flipH,attr,omitempty"` // Horizontal flip
	FlipV *bool `xml:"flipV,attr,omitempty"` // Vertical flip
	Off   *Off  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main off,omitempty"`
	Ext   *Ext  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ext,omitempty"`
}

// Off represents an offset.
type Off struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// Ext represents extents.
type Ext struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

// PrstGeom represents preset geometry.
type PrstGeom struct {
	Prst  string  `xml:"prst,attr"` // rect, ellipse, roundRect, etc.
	AvLst *AvLst  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main avLst,omitempty"`
}

// AvLst represents adjustment values list.
type AvLst struct {
	Gd []*Gd `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gd,omitempty"`
}

// Gd represents a guide (adjustment value).
type Gd struct {
	Name string `xml:"name,attr"`
	Fmla string `xml:"fmla,attr"`
}

// CustGeom represents custom geometry.
type CustGeom struct {
	// Custom geometry elements
}

// ShapeStyle represents shape style.
type ShapeStyle struct {
	LnRef     *StyleRef `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lnRef,omitempty"`
	FillRef   *StyleRef `xml:"http://schemas.openxmlformats.org/drawingml/2006/main fillRef,omitempty"`
	EffectRef *StyleRef `xml:"http://schemas.openxmlformats.org/drawingml/2006/main effectRef,omitempty"`
	FontRef   *FontRef  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main fontRef,omitempty"`
}

// StyleRef represents a style matrix reference.
type StyleRef struct {
	Idx       int        `xml:"idx,attr"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

// FontRef represents a font reference.
type FontRef struct {
	Idx       string     `xml:"idx,attr"` // major, minor
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

// Common preset geometry types.
const (
	PrstGeomRect              = "rect"
	PrstGeomEllipse           = "ellipse"
	PrstGeomRoundRect         = "roundRect"
	PrstGeomTriangle          = "triangle"
	PrstGeomRightArrow        = "rightArrow"
	PrstGeomLeftArrow         = "leftArrow"
	PrstGeomUpArrow           = "upArrow"
	PrstGeomDownArrow         = "downArrow"
	PrstGeomDiamond           = "diamond"
	PrstGeomPentagon          = "pentagon"
	PrstGeomHexagon           = "hexagon"
	PrstGeomStar5             = "star5"
	PrstGeomStar6             = "star6"
	PrstGeomLine              = "line"
	PrstGeomBentConnector3    = "bentConnector3"
	PrstGeomCurvedConnector3  = "curvedConnector3"
)
