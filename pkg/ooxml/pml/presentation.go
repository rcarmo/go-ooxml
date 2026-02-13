// Package pml provides PresentationML types for OOXML presentations.
package pml

import (
	"encoding/xml"
	"fmt"
)

// Namespaces used in PresentationML documents.
const (
	NS  = "http://schemas.openxmlformats.org/presentationml/2006/main"
	NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	NSA = "http://schemas.openxmlformats.org/drawingml/2006/main"
)

// Presentation represents the presentation part.
type Presentation struct {
	XMLName            xml.Name            `xml:"http://schemas.openxmlformats.org/presentationml/2006/main presentation"`
	XMLNS_R            string              `xml:"xmlns:r,attr,omitempty"`
	SldMasterIdLst     *SldMasterIdLst     `xml:"sldMasterIdLst,omitempty"`
	NotesMasterIdLst   *NotesMasterIdLst   `xml:"notesMasterIdLst,omitempty"`
	SldIdLst           *SldIdLst           `xml:"sldIdLst,omitempty"`
	SldSz              *SldSz              `xml:"sldSz,omitempty"`
	NotesSz            *NotesSz            `xml:"notesSz,omitempty"`
	DefaultTextStyle   *DefaultTextStyle   `xml:"defaultTextStyle,omitempty"`
	ExtLst             *ExtLst             `xml:"extLst,omitempty"`
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

// NotesMasterIdLst is a list of notes master IDs.
type NotesMasterIdLst struct {
	NotesMasterId []*NotesMasterId `xml:"notesMasterId,omitempty"`
}

// NotesMasterId references a notes master.
type NotesMasterId struct {
	RID string `xml:"-"` // Manually handled
}

// UnmarshalXML customizes XML parsing to properly read r:id attribute.
func (n *NotesMasterId) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Space != "" && attr.Name.Local == "id" {
			n.RID = attr.Value
			break
		}
	}
	return d.Skip()
}

// MarshalXML customizes XML output to use r: prefix for relationship ID.
func (n NotesMasterId) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "notesMasterId"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Space: NSR, Local: "id"}, Value: n.RID},
	}
	return e.EncodeElement("", start)
}

// UnmarshalXML customizes XML parsing to properly read r:id attribute.
func (s *SldMasterId) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			if attr.Name.Space == "" {
				var id int
				if _, err := fmt.Sscanf(attr.Value, "%d", &id); err == nil {
					s.ID = id
				}
			} else {
				s.RID = attr.Value
			}
		}
	}
	return d.Skip()
}

// MarshalXML customizes XML output to use r: prefix for relationship ID.
func (s SldMasterId) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "sldMasterId"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "id"}, Value: fmt.Sprintf("%d", s.ID)},
		{Name: xml.Name{Space: NSR, Local: "id"}, Value: s.RID},
	}
	return e.EncodeElement("", start)
}

// SldIdLst is a list of slide IDs.
type SldIdLst struct {
	SldId []*SldId `xml:"sldId,omitempty"`
}

// SldId references a slide.
type SldId struct {
	ID  int    `xml:"id,attr"`
	RID string `xml:"-"` // Manually handled
}

// UnmarshalXML customizes XML parsing to properly read r:id attribute.
func (s *SldId) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			if attr.Name.Space == "" {
				// This is the numeric id
				var id int
				_, err := fmt.Sscanf(attr.Value, "%d", &id)
				if err == nil {
					s.ID = id
				}
			} else {
				// This is the relationship id (r:id)
				s.RID = attr.Value
			}
		}
	}
	// Skip any content
	return d.Skip()
}

// MarshalXML customizes XML output to use r: prefix for relationship ID.
func (s SldId) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "sldId"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "id"}, Value: fmt.Sprintf("%d", s.ID)},
		{Name: xml.Name{Space: NSR, Local: "id"}, Value: s.RID},
	}
	return e.EncodeElement("", start)
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
