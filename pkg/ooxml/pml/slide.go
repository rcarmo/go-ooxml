package pml

import "encoding/xml"

// Sld represents a slide.
type Sld struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/presentationml/2006/main sld"`
	CSld    *CSld    `xml:"cSld"`
	ClrMapOvr *ClrMapOvr `xml:"clrMapOvr,omitempty"`
	Timing  *Timing  `xml:"timing,omitempty"`
	Show    *bool    `xml:"show,attr,omitempty"` // false means hidden
}

// CSld represents common slide data.
type CSld struct {
	Name   string  `xml:"name,attr,omitempty"`
	SpTree *SpTree `xml:"spTree"`
	Bg     *Bg     `xml:"bg,omitempty"`
}

// SpTree is a shape tree containing all shapes on the slide.
type SpTree struct {
	NvGrpSpPr *NvGrpSpPr `xml:"nvGrpSpPr"`
	GrpSpPr   *GrpSpPr   `xml:"grpSpPr"`
	Content   []interface{} `xml:",any"` // Sp, Pic, GraphicFrame, etc.
}

// NvGrpSpPr represents non-visual group shape properties.
type NvGrpSpPr struct {
	CNvPr      *CNvPr      `xml:"cNvPr"`
	CNvGrpSpPr *CNvGrpSpPr `xml:"cNvGrpSpPr"`
	NvPr       *NvPr       `xml:"nvPr"`
}

// CNvPr represents common non-visual properties.
type CNvPr struct {
	ID    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Descr string `xml:"descr,attr,omitempty"`
	Title string `xml:"title,attr,omitempty"`
	Hidden *bool `xml:"hidden,attr,omitempty"`
}

// CNvGrpSpPr represents non-visual group shape drawing properties.
type CNvGrpSpPr struct {
}

// NvPr represents non-visual properties.
type NvPr struct {
	IsPhoto *bool  `xml:"isPhoto,attr,omitempty"`
	Ph      *Ph    `xml:"ph,omitempty"`
}

// Ph represents placeholder information.
type Ph struct {
	Type string `xml:"type,attr,omitempty"` // title, body, ctrTitle, subTitle, etc.
	Idx  *int   `xml:"idx,attr,omitempty"`
	Sz   string `xml:"sz,attr,omitempty"`   // full, half, quarter
	Orient string `xml:"orient,attr,omitempty"` // horz, vert
	HasCustomPrompt *bool `xml:"hasCustomPrompt,attr,omitempty"`
}

// GrpSpPr represents group shape properties.
type GrpSpPr struct {
	Xfrm *GrpXfrm `xml:"http://schemas.openxmlformats.org/drawingml/2006/main xfrm,omitempty"`
}

// GrpXfrm represents group transform.
type GrpXfrm struct {
	Off     *Off     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main off,omitempty"`
	Ext     *Ext     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ext,omitempty"`
	ChOff   *Off     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main chOff,omitempty"`
	ChExt   *Ext     `xml:"http://schemas.openxmlformats.org/drawingml/2006/main chExt,omitempty"`
}

// Off represents an offset (position).
type Off struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// Ext represents extents (size).
type Ext struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

// Bg represents background.
type Bg struct {
	BgPr *BgPr `xml:"bgPr,omitempty"`
}

// BgPr represents background properties.
type BgPr struct {
	// Background fill (solid, gradient, blip, etc.)
}

// ClrMapOvr represents color map override.
type ClrMapOvr struct {
	MasterClrMapping *MasterClrMapping `xml:"masterClrMapping,omitempty"`
}

// MasterClrMapping indicates to use master color mapping.
type MasterClrMapping struct {
}

// Timing represents slide timing.
type Timing struct {
	TnLst *TnLst `xml:"tnLst,omitempty"`
}

// TnLst is a list of time nodes.
type TnLst struct {
}

// Placeholder type constants.
const (
	PhTypeTitle     = "title"
	PhTypeCtrTitle  = "ctrTitle"
	PhTypeSubTitle  = "subTitle"
	PhTypeBody      = "body"
	PhTypeDT        = "dt"       // Date/time
	PhTypeFtr       = "ftr"      // Footer
	PhTypeSldNum    = "sldNum"   // Slide number
	PhTypeTbl       = "tbl"      // Table
	PhTypeChart     = "chart"
	PhTypeDgm       = "dgm"      // Diagram
	PhTypeMedia     = "media"
	PhTypeClipArt   = "clipArt"
	PhTypePic       = "pic"      // Picture
	PhTypeObj       = "obj"      // Object
	PhTypeSldImg    = "sldImg"   // Slide image
)
