package wml

import "encoding/xml"

// Styles represents the styles part.
type Styles struct {
	XMLName      xml.Name    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main styles"`
	DocDefaults  *DocDefaults `xml:"docDefaults,omitempty"`
	LatentStyles *LatentStyles `xml:"latentStyles,omitempty"`
	Styles       []*Style     `xml:"style,omitempty"`
}

// DocDefaults represents document defaults.
type DocDefaults struct {
	RPrDefault *RPrDefault `xml:"rPrDefault,omitempty"`
	PPrDefault *PPrDefault `xml:"pPrDefault,omitempty"`
}

// RPrDefault represents default run properties.
type RPrDefault struct {
	RPr *RPr `xml:"rPr,omitempty"`
}

// PPrDefault represents default paragraph properties.
type PPrDefault struct {
	PPr *PPr `xml:"pPr,omitempty"`
}

// LatentStyles represents latent style definitions.
type LatentStyles struct {
	DefLockedState    *bool            `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main defLockedState,attr,omitempty"`
	DefUIPriority     int              `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main defUIPriority,attr,omitempty"`
	DefSemiHidden     *bool            `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main defSemiHidden,attr,omitempty"`
	DefUnhideWhenUsed *bool            `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main defUnhideWhenUsed,attr,omitempty"`
	DefQFormat        *bool            `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main defQFormat,attr,omitempty"`
	Count             int              `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main count,attr,omitempty"`
	LsdException      []*LsdException  `xml:"lsdException,omitempty"`
}

// LsdException represents a latent style exception.
type LsdException struct {
	Name          string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main name,attr"`
	SemiHidden    *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main semiHidden,attr,omitempty"`
	UIPriority    *int   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main uiPriority,attr,omitempty"`
	UnhideWhenUsed *bool `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main unhideWhenUsed,attr,omitempty"`
	QFormat       *bool  `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main qFormat,attr,omitempty"`
}

// Style represents a style definition.
type Style struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main style"`
	Type        string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main type,attr,omitempty"`
	StyleID     string   `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main styleId,attr,omitempty"`
	Default     *bool    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main default,attr,omitempty"`
	CustomStyle *bool    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main customStyle,attr,omitempty"`
	Name        *StyleName `xml:"name,omitempty"`
	Aliases     *StyleAliases `xml:"aliases,omitempty"`
	BasedOn     *StyleBasedOn `xml:"basedOn,omitempty"`
	Next        *StyleNext `xml:"next,omitempty"`
	Link        *StyleLink `xml:"link,omitempty"`
	AutoRedefine *OnOff `xml:"autoRedefine,omitempty"`
	Hidden      *OnOff `xml:"hidden,omitempty"`
	UIPriority  *UIPriority `xml:"uiPriority,omitempty"`
	SemiHidden  *OnOff `xml:"semiHidden,omitempty"`
	UnhideWhenUsed *OnOff `xml:"unhideWhenUsed,omitempty"`
	QFormat     *OnOff `xml:"qFormat,omitempty"`
	Locked      *OnOff `xml:"locked,omitempty"`
	Personal    *OnOff `xml:"personal,omitempty"`
	PersonalCompose *OnOff `xml:"personalCompose,omitempty"`
	PersonalReply *OnOff `xml:"personalReply,omitempty"`
	Rsid        *Rsid  `xml:"rsid,omitempty"`
	PPr         *PPr   `xml:"pPr,omitempty"`
	RPr         *RPr   `xml:"rPr,omitempty"`
	TblPr       *TblPr `xml:"tblPr,omitempty"`
	TrPr        *TrPr  `xml:"trPr,omitempty"`
	TcPr        *TcPr  `xml:"tcPr,omitempty"`
}

// StyleName represents a style name.
type StyleName struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// StyleAliases represents style aliases.
type StyleAliases struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// StyleBasedOn represents the base style.
type StyleBasedOn struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// StyleNext represents the next style.
type StyleNext struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// StyleLink represents a linked style.
type StyleLink struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// UIPriority represents UI priority.
type UIPriority struct {
	Val int `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// StyleType constants.
const (
	StyleTypeParagraph = "paragraph"
	StyleTypeCharacter = "character"
	StyleTypeTable     = "table"
	StyleTypeNumbering = "numbering"
)
