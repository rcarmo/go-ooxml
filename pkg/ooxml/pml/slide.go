package pml

import (
	"encoding/xml"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
)

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

// UnmarshalXML implements custom XML unmarshaling for SpTree to capture shapes.
func (s *SpTree) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "nvGrpSpPr":
				s.NvGrpSpPr = &NvGrpSpPr{}
				if err := d.DecodeElement(s.NvGrpSpPr, &t); err != nil {
					return err
				}
			case "grpSpPr":
				s.GrpSpPr = &GrpSpPr{}
				if err := d.DecodeElement(s.GrpSpPr, &t); err != nil {
					return err
				}
			case "sp":
				sp := &dml.Sp{}
				if err := d.DecodeElement(sp, &t); err != nil {
					return err
				}
				s.Content = append(s.Content, sp)
			case "graphicFrame":
				gf := &GraphicFrame{}
				if err := d.DecodeElement(gf, &t); err != nil {
					return err
				}
				s.Content = append(s.Content, gf)
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if t.Name == start.Name {
				return nil
			}
		}
	}
}

// MarshalXML implements custom XML marshaling for SpTree.
func (s *SpTree) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "spTree"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if s.NvGrpSpPr != nil {
		if err := e.EncodeElement(s.NvGrpSpPr, xml.StartElement{Name: xml.Name{Space: NS, Local: "nvGrpSpPr"}}); err != nil {
			return err
		}
	}
	if s.GrpSpPr != nil {
		if err := e.EncodeElement(s.GrpSpPr, xml.StartElement{Name: xml.Name{Space: NS, Local: "grpSpPr"}}); err != nil {
			return err
		}
	}
	for _, elem := range s.Content {
		switch v := elem.(type) {
		case *dml.Sp:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "sp"}}); err != nil {
				return err
			}
		case *GraphicFrame:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "graphicFrame"}}); err != nil {
				return err
			}
		default:
			if err := e.Encode(v); err != nil {
				return err
			}
		}
	}
	return e.EncodeToken(start.End())
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

// GraphicFrame represents a graphic frame (e.g., table).
type GraphicFrame struct {
	XMLName           xml.Name         `xml:"http://schemas.openxmlformats.org/presentationml/2006/main graphicFrame"`
	NvGraphicFramePr  *NvGraphicFramePr `xml:"nvGraphicFramePr"`
	Xfrm              *Xfrm             `xml:"xfrm"`
	Graphic           *dml.Graphic      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main graphic"`
}

// NvGraphicFramePr represents non-visual graphic frame properties.
type NvGraphicFramePr struct {
	CNvPr   *CNvPr   `xml:"cNvPr"`
	CNvGraphicFramePr *CNvGraphicFramePr `xml:"cNvGraphicFramePr"`
	NvPr    *NvPr    `xml:"nvPr"`
}

// CNvGraphicFramePr represents non-visual graphic frame drawing properties.
type CNvGraphicFramePr struct{}

// Xfrm represents a transform for graphic frames.
type Xfrm struct {
	Off *Off `xml:"http://schemas.openxmlformats.org/drawingml/2006/main off,omitempty"`
	Ext *Ext `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ext,omitempty"`
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
