package wml

import "encoding/xml"

// R represents a text run.
type R struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main r"`
	RPr     *RPr     `xml:"rPr,omitempty"`
	Content []interface{} `xml:"-"` // Handled by custom unmarshaler
}

// UnmarshalXML implements custom XML unmarshaling for R.
func (r *R) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPr":
				r.RPr = &RPr{}
				if err := d.DecodeElement(r.RPr, &t); err != nil {
					return err
				}
			case "t":
				text := &T{}
				if err := d.DecodeElement(text, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, text)
			case "br":
				br := &Br{}
				if err := d.DecodeElement(br, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, br)
			case "tab":
				tab := &Tab{}
				if err := d.DecodeElement(tab, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, tab)
			case "delText":
				dt := &DelText{}
				if err := d.DecodeElement(dt, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, dt)
			case "fldChar":
				fc := &FldChar{}
				if err := d.DecodeElement(fc, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, fc)
		case "instrText":
			it := &InstrText{}
			if err := d.DecodeElement(it, &t); err != nil {
				return err
			}
			r.Content = append(r.Content, it)
		case "drawing":
			drawing := &Drawing{}
			if err := d.DecodeElement(drawing, &t); err != nil {
				return err
			}
			r.Content = append(r.Content, drawing)
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

// MarshalXML implements custom XML marshaling for R.
func (r *R) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: NS, Local: "r"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if r.RPr != nil {
		if err := e.Encode(r.RPr); err != nil {
			return err
		}
	}

	for _, elem := range r.Content {
		if err := e.Encode(elem); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// RPr represents run properties.
type RPr struct {
	XMLName      xml.Name      `xml:"rPr"`
	RStyle       *RStyle       `xml:"rStyle,omitempty"`
	B            *OnOff        `xml:"b,omitempty"`
	BCs          *OnOff        `xml:"bCs,omitempty"`
	I            *OnOff        `xml:"i,omitempty"`
	ICs          *OnOff        `xml:"iCs,omitempty"`
	Caps         *OnOff        `xml:"caps,omitempty"`
	SmallCaps    *OnOff        `xml:"smallCaps,omitempty"`
	Strike       *OnOff        `xml:"strike,omitempty"`
	Dstrike      *OnOff        `xml:"dstrike,omitempty"`
	Outline      *OnOff        `xml:"outline,omitempty"`
	Shadow       *OnOff        `xml:"shadow,omitempty"`
	Emboss       *OnOff        `xml:"emboss,omitempty"`
	Imprint      *OnOff        `xml:"imprint,omitempty"`
	NoProof      *OnOff        `xml:"noProof,omitempty"`
	SnapToGrid   *OnOff        `xml:"snapToGrid,omitempty"`
	Vanish       *OnOff        `xml:"vanish,omitempty"`
	Color        *Color        `xml:"color,omitempty"`
	Spacing      *RPrSpacing   `xml:"spacing,omitempty"`
	W            *RPrW         `xml:"w,omitempty"`
	Kern         *Kern         `xml:"kern,omitempty"`
	Position     *Position     `xml:"position,omitempty"`
	Sz           *Sz           `xml:"sz,omitempty"`
	SzCs         *Sz           `xml:"szCs,omitempty"`
	Highlight    *Highlight    `xml:"highlight,omitempty"`
	U            *U            `xml:"u,omitempty"`
	Effect       *Effect       `xml:"effect,omitempty"`
	RFonts       *RFonts       `xml:"rFonts,omitempty"`
	VertAlign    *VertAlign    `xml:"vertAlign,omitempty"`
	Lang         *Lang         `xml:"lang,omitempty"`
}

// RStyle references a character style.
type RStyle struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Color represents text color.
type Color struct {
	Val       string  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	ThemeColor string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main themeColor,attr,omitempty"`
	ThemeTint  string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main themeTint,attr,omitempty"`
	ThemeShade string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main themeShade,attr,omitempty"`
}

// RPrSpacing represents character spacing.
type RPrSpacing struct {
	Val int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// RPrW represents character width.
type RPrW struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Kern represents kerning.
type Kern struct {
	Val int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Position represents vertical position.
type Position struct {
	Val int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Sz represents font size in half-points.
type Sz struct {
	Val int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Highlight represents text highlighting.
type Highlight struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// U represents underlining.
type U struct {
	Val   string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	Color string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main color,attr,omitempty"`
}

// Effect represents text effects.
type Effect struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// RFonts represents font settings.
type RFonts struct {
	Ascii    string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main ascii,attr,omitempty"`
	HAnsi    string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main hAnsi,attr,omitempty"`
	EastAsia string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main eastAsia,attr,omitempty"`
	Cs       string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main cs,attr,omitempty"`
}

// VertAlign represents vertical alignment (subscript/superscript).
type VertAlign struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Lang represents language settings.
type Lang struct {
	Val      string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
	EastAsia string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main eastAsia,attr,omitempty"`
	Bidi     string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main bidi,attr,omitempty"`
}

// T represents text content.
type T struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main t"`
	Space   string   `xml:"http://www.w3.org/XML/1998/namespace space,attr,omitempty"`
	Text    string   `xml:",chardata"`
}

// Br represents a break.
type Br struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main br"`
	Type    string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main type,attr,omitempty"`
}

// Tab represents a tab character.
type Tab struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main tab"`
}

// NewT creates a new text element.
func NewT(text string) *T {
	t := &T{Text: text}
	// Preserve whitespace if needed
	if len(text) > 0 && (text[0] == ' ' || text[len(text)-1] == ' ' || containsMultipleSpaces(text)) {
		t.Space = "preserve"
	}
	return t
}

func containsMultipleSpaces(s string) bool {
	prev := false
	for _, c := range s {
		if c == ' ' {
			if prev {
				return true
			}
			prev = true
		} else {
			prev = false
		}
	}
	return false
}
