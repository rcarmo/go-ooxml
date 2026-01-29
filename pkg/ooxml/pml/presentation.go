// Package pml provides PresentationML types for OOXML presentations.
package pml

import "encoding/xml"

// Namespaces used in PresentationML documents.
const (
	NS  = "http://schemas.openxmlformats.org/presentationml/2006/main"
	NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	NSA = "http://schemas.openxmlformats.org/drawingml/2006/main"
)

// Presentation represents the presentation part.
type Presentation struct {
	XMLName            xml.Name            `xml:"http://schemas.openxmlformats.org/presentationml/2006/main presentation"`
	SldMasterIdLst     *SldMasterIdLst     `xml:"sldMasterIdLst,omitempty"`
	SldIdLst           *SldIdLst           `xml:"sldIdLst,omitempty"`
	SldSz              *SldSz              `xml:"sldSz,omitempty"`
	NotesSz            *NotesSz            `xml:"notesSz,omitempty"`
	DefaultTextStyle   *DefaultTextStyle   `xml:"defaultTextStyle,omitempty"`
}

// SldMasterIdLst is a list of slide master IDs.
type SldMasterIdLst struct {
	SldMasterId []*SldMasterId `xml:"sldMasterId,omitempty"`
}

// SldMasterId references a slide master.
type SldMasterId struct {
	ID  int    `xml:"id,attr,omitempty"`
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// SldIdLst is a list of slide IDs.
type SldIdLst struct {
	SldId []*SldId `xml:"sldId,omitempty"`
}

// SldId references a slide.
type SldId struct {
	ID  int    `xml:"id,attr"`
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// SldSz represents slide size.
type SldSz struct {
	Cx   int64  `xml:"cx,attr"` // Width in EMUs
	Cy   int64  `xml:"cy,attr"` // Height in EMUs
	Type string `xml:"type,attr,omitempty"` // screen4x3, screen16x9, screen16x10, etc.
}

// NotesSz represents notes slide size.
type NotesSz struct {
	Cx int64 `xml:"cx,attr"` // Width in EMUs
	Cy int64 `xml:"cy,attr"` // Height in EMUs
}

// DefaultTextStyle represents default text styling.
type DefaultTextStyle struct {
	DefPPr *DefPPr   `xml:"defPPr,omitempty"`
	LvlPPr []*LvlPPr `xml:"lvl1pPr,omitempty"` // Actually lvl1pPr through lvl9pPr
}

// DefPPr represents default paragraph properties.
type DefPPr struct {
	// Text paragraph properties from DML
}

// LvlPPr represents level-specific paragraph properties.
type LvlPPr struct {
	// Level paragraph properties from DML
}

// Common slide size types.
const (
	SldSzScreen4x3   = "screen4x3"
	SldSzScreen16x9  = "screen16x9"
	SldSzScreen16x10 = "screen16x10"
	SldSzLetter      = "letter"
	SldSzLedger      = "ledger"
	SldSzA3          = "A3"
	SldSzA4          = "A4"
	SldSzB4ISO       = "B4ISO"
	SldSzB5ISO       = "B5ISO"
	SldSzB4JIS       = "B4JIS"
	SldSzB5JIS       = "B5JIS"
	SldSz35mm        = "35mm"
	SldSzOverhead    = "overhead"
	SldSzBanner      = "banner"
	SldSzCustom      = "custom"
)
