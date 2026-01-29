package wml

import "encoding/xml"

// P represents a paragraph.
type P struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main p"`
	PPr     *PPr     `xml:"pPr,omitempty"`
	Content []interface{} `xml:"-"` // Handled by custom unmarshaler
}

// UnmarshalXML implements custom XML unmarshaling for P.
func (p *P) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pPr":
				p.PPr = &PPr{}
				if err := d.DecodeElement(p.PPr, &t); err != nil {
					return err
				}
			case "r":
				r := &R{}
				if err := d.DecodeElement(r, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, r)
			case "ins":
				ins := &Ins{}
				if err := d.DecodeElement(ins, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, ins)
			case "del":
				del := &Del{}
				if err := d.DecodeElement(del, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, del)
			case "commentRangeStart":
				crs := &CommentRangeStart{}
				if err := d.DecodeElement(crs, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, crs)
			case "commentRangeEnd":
				cre := &CommentRangeEnd{}
				if err := d.DecodeElement(cre, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, cre)
			case "bookmarkStart":
				bs := &BookmarkStart{}
				if err := d.DecodeElement(bs, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, bs)
			case "bookmarkEnd":
				be := &BookmarkEnd{}
				if err := d.DecodeElement(be, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, be)
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

// MarshalXML implements custom XML marshaling for P.
func (p *P) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "p"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if p.PPr != nil {
		if err := e.Encode(p.PPr); err != nil {
			return err
		}
	}

	for _, elem := range p.Content {
		if err := e.Encode(elem); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// BookmarkStart represents the start of a bookmark.
type BookmarkStart struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main bookmarkStart"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
	Name    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main name,attr"`
}

// BookmarkEnd represents the end of a bookmark.
type BookmarkEnd struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main bookmarkEnd"`
	ID      int      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main id,attr"`
}

// PPr represents paragraph properties.
type PPr struct {
	XMLName    xml.Name    `xml:"pPr"`
	PStyle     *PStyle     `xml:"pStyle,omitempty"`
	KeepNext   *OnOff      `xml:"keepNext,omitempty"`
	KeepLines  *OnOff      `xml:"keepLines,omitempty"`
	PageBreakBefore *OnOff `xml:"pageBreakBefore,omitempty"`
	Spacing    *Spacing    `xml:"spacing,omitempty"`
	Ind        *Ind        `xml:"ind,omitempty"`
	Jc         *Jc         `xml:"jc,omitempty"`
	RPr        *RPr        `xml:"rPr,omitempty"`
	NumPr      *NumPr      `xml:"numPr,omitempty"`
	OutlineLvl *OutlineLvl `xml:"outlineLvl,omitempty"`
}

// PStyle references a paragraph style.
type PStyle struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// OnOff represents a boolean toggle.
type OnOff struct {
	Val *bool `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
}

// Spacing represents paragraph spacing.
type Spacing struct {
	Before   *int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main before,attr,omitempty"`
	After    *int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main after,attr,omitempty"`
	Line     *int64  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main line,attr,omitempty"`
	LineRule *string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main lineRule,attr,omitempty"`
}

// Ind represents paragraph indentation.
type Ind struct {
	Left      *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main left,attr,omitempty"`
	Right     *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main right,attr,omitempty"`
	FirstLine *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main firstLine,attr,omitempty"`
	Hanging   *int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hanging,attr,omitempty"`
}

// Jc represents paragraph justification.
type Jc struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// NumPr represents numbering properties.
type NumPr struct {
	Ilvl  *Ilvl  `xml:"ilvl,omitempty"`
	NumID *NumID `xml:"numId,omitempty"`
}

// Ilvl represents indentation level for numbering.
type Ilvl struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// NumID references a numbering definition.
type NumID struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// OutlineLvl represents outline level.
type OutlineLvl struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Enabled returns true if the OnOff element is set and enabled.
func (o *OnOff) Enabled() bool {
	if o == nil {
		return false
	}
	if o.Val == nil {
		return true // presence without val means true
	}
	return *o.Val
}

// NewOnOff creates an OnOff element with the given value.
func NewOnOff(val bool) *OnOff {
	return &OnOff{Val: &val}
}

// NewOnOffEnabled creates an enabled OnOff element (no val attribute).
func NewOnOffEnabled() *OnOff {
	return &OnOff{}
}
