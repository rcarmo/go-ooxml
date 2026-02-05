package pml

import (
	"encoding/xml"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
)

// Pic represents a picture in a slide.
type Pic struct {
	XMLName  xml.Name  `xml:"http://schemas.openxmlformats.org/presentationml/2006/main pic"`
	NvPicPr  *NvPicPr  `xml:"nvPicPr"`
	BlipFill *dml.BlipFill `xml:"http://schemas.openxmlformats.org/drawingml/2006/main blipFill"`
	SpPr     *dml.SpPr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spPr"`
}

// NvPicPr represents non-visual picture properties.
type NvPicPr struct {
	CNvPr    *CNvPr    `xml:"cNvPr"`
	CNvPicPr *CNvPicPr `xml:"cNvPicPr"`
	NvPr     *NvPr     `xml:"nvPr"`
}

// CNvPicPr represents non-visual picture drawing properties.
type CNvPicPr struct {
	PicLocks *PicLocks `xml:"http://schemas.openxmlformats.org/drawingml/2006/main picLocks,omitempty"`
}

// PicLocks represents picture locking settings.
type PicLocks struct {
	NoChangeAspect   *bool `xml:"noChangeAspect,attr,omitempty"`
	NoChangeArrowheads *bool `xml:"noChangeArrowheads,attr,omitempty"`
	NoCrop           *bool `xml:"noCrop,attr,omitempty"`
	NoMove           *bool `xml:"noMove,attr,omitempty"`
	NoResize         *bool `xml:"noResize,attr,omitempty"`
	NoRot            *bool `xml:"noRot,attr,omitempty"`
}
