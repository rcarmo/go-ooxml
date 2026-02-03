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
			case "hyperlink":
				h := &Hyperlink{}
				if err := d.DecodeElement(h, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, h)
			case "sdt":
				sdt := &Sdt{}
				if err := d.DecodeElement(sdt, &t); err != nil {
					return err
				}
				p.Content = append(p.Content, sdt)
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

// Hyperlink represents a hyperlink in a paragraph.
type Hyperlink struct {
	XMLName xml.Name      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hyperlink"`
	ID      string        `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	Anchor  string        `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main anchor,attr,omitempty"`
	Tooltip string        `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tooltip,attr,omitempty"`
	History *bool         `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main history,attr,omitempty"`
	Content []interface{} `xml:"-"` // runs and other content
}

// UnmarshalXML implements custom XML unmarshaling for Hyperlink.
func (h *Hyperlink) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	h.XMLName = start.Name
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			h.ID = attr.Value
		case "anchor":
			h.Anchor = attr.Value
		case "tooltip":
			h.Tooltip = attr.Value
		case "history":
			val := attr.Value == "1" || attr.Value == "true"
			h.History = &val
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "r":
				r := &R{}
				if err := d.DecodeElement(r, &t); err != nil {
					return err
				}
				h.Content = append(h.Content, r)
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

// MarshalXML implements custom XML marshaling for Hyperlink.
func (h *Hyperlink) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "hyperlink"}
	if h.ID != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: NSR, Local: "id"}, Value: h.ID})
	}
	if h.Anchor != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: NS, Local: "anchor"}, Value: h.Anchor})
	}
	if h.Tooltip != "" {
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: NS, Local: "tooltip"}, Value: h.Tooltip})
	}
	if h.History != nil {
		val := "0"
		if *h.History {
			val = "1"
		}
		start.Attr = append(start.Attr, xml.Attr{Name: xml.Name{Space: NS, Local: "history"}, Value: val})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, elem := range h.Content {
		if err := e.Encode(elem); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// Sdt represents a content control.
type Sdt struct {
	XMLName    xml.Name    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sdt"`
	SdtPr      *SdtPr      `xml:"sdtPr,omitempty"`
	SdtContent *SdtContent `xml:"sdtContent,omitempty"`
}

// SdtPr represents content control properties.
type SdtPr struct {
	Alias         *SdtString        `xml:"alias,omitempty"`
	Tag           *SdtString        `xml:"tag,omitempty"`
	ID            *SdtID            `xml:"id,omitempty"`
	ShowingPlcHdr *OnOff            `xml:"showingPlcHdr,omitempty"`
	Lock          *SdtLock          `xml:"lock,omitempty"`
	DropDownList  *SdtDropDownList  `xml:"dropDownList,omitempty"`
	ComboBox      *SdtDropDownList  `xml:"comboBox,omitempty"`
	Date          *SdtDate          `xml:"date,omitempty"`
}

// SdtContent represents content control contents.
type SdtContent struct {
	XMLName xml.Name      `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main sdtContent"`
	Content []interface{} `xml:"-"` // paragraphs, runs, tables
}

// UnmarshalXML implements custom XML unmarshaling for SdtContent.
func (c *SdtContent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "p":
				p := &P{}
				if err := d.DecodeElement(p, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, p)
			case "tbl":
				tbl := &Tbl{}
				if err := d.DecodeElement(tbl, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, tbl)
			case "r":
				r := &R{}
				if err := d.DecodeElement(r, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, r)
			case "hyperlink":
				h := &Hyperlink{}
				if err := d.DecodeElement(h, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, h)
			case "sdt":
				sdt := &Sdt{}
				if err := d.DecodeElement(sdt, &t); err != nil {
					return err
				}
				c.Content = append(c.Content, sdt)
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

// MarshalXML implements custom XML marshaling for SdtContent.
func (c *SdtContent) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "sdtContent"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, elem := range c.Content {
		switch v := elem.(type) {
		case *P:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "p"}}); err != nil {
				return err
			}
		case *Tbl:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "tbl"}}); err != nil {
				return err
			}
		case *R:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "r"}}); err != nil {
				return err
			}
		case *Hyperlink:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "hyperlink"}}); err != nil {
				return err
			}
		case *Sdt:
			if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: NS, Local: "sdt"}}); err != nil {
				return err
			}
		default:
			if err := e.Encode(elem); err != nil {
				return err
			}
		}
	}
	return e.EncodeToken(start.End())
}

// SdtString represents a string property for content controls.
type SdtString struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// SdtID represents a content control ID.
type SdtID struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// SdtLock represents a content control lock setting.
type SdtLock struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// SdtDropDownList represents dropdown/combo box entries.
type SdtDropDownList struct {
	ListItem []*SdtListItem `xml:"listItem,omitempty"`
}

// SdtListItem represents a dropdown list item.
type SdtListItem struct {
	DisplayText string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main displayText,attr,omitempty"`
	Value       string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main value,attr,omitempty"`
}

// SdtDate represents date picker properties.
type SdtDate struct {
	DateFormat        *SdtString    `xml:"dateFormat,omitempty"`
	Language          *SdtLang      `xml:"lid,omitempty"`
	StoreMappedDataAs *SdtString    `xml:"storeMappedDataAs,omitempty"`
	Calendar          *SdtString    `xml:"calendar,omitempty"`
	FullDate          *SdtDateValue `xml:"fullDate,omitempty"`
}

// SdtLang represents a language identifier.
type SdtLang struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// SdtDateValue represents a full date value.
type SdtDateValue struct {
	Val string `xml:"http://schemas.microsoft.com/office/word/2010/wordml val,attr"`
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
