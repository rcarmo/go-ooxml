package wml

import "encoding/xml"

// Settings represents document settings.
type Settings struct {
	XMLName            xml.Name            `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main settings"`
	Zoom               *Zoom               `xml:"zoom,omitempty"`
	TrackRevisions     *OnOff              `xml:"trackRevisions,omitempty"`
	DefaultTabStop     *DefaultTabStop     `xml:"defaultTabStop,omitempty"`
	CharacterSpacingControl *CharacterSpacingControl `xml:"characterSpacingControl,omitempty"`
	Compat             *Compat             `xml:"compat,omitempty"`
	Rsids              *Rsids              `xml:"rsids,omitempty"`
}

// Zoom represents document zoom settings.
type Zoom struct {
	Percent int    `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main percent,attr,omitempty"`
	Val     string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
}

// DefaultTabStop represents the default tab stop value.
type DefaultTabStop struct {
	Val int64 `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// CharacterSpacingControl represents character spacing control.
type CharacterSpacingControl struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Compat represents compatibility settings.
type Compat struct {
	CompatSetting []*CompatSetting `xml:"compatSetting,omitempty"`
}

// CompatSetting represents a compatibility setting.
type CompatSetting struct {
	Name string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main name,attr,omitempty"`
	URI  string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main uri,attr,omitempty"`
	Val  string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr,omitempty"`
}

// Rsids represents revision save IDs.
type Rsids struct {
	RsidRoot *RsidRoot `xml:"rsidRoot,omitempty"`
	Rsid     []*Rsid   `xml:"rsid,omitempty"`
}

// RsidRoot represents the root RSID.
type RsidRoot struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}

// Rsid represents a revision save ID.
type Rsid struct {
	Val string `xml:"http://schemas.openxmlformats.org/wordprocessingml/2006/main val,attr"`
}
