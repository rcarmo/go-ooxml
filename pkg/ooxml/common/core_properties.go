package common

import "encoding/xml"

// CoreProperties represents document core properties (Dublin Core).
type CoreProperties struct {
	XMLName        xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties coreProperties"`
	Category       string   `xml:"http://purl.org/dc/elements/1.1/ category,omitempty"`
	ContentStatus  string   `xml:"contentStatus,omitempty"`
	Created        *DCDate  `xml:"http://purl.org/dc/terms/ created,omitempty"`
	Creator        string   `xml:"http://purl.org/dc/elements/1.1/ creator,omitempty"`
	Description    string   `xml:"http://purl.org/dc/elements/1.1/ description,omitempty"`
	Identifier     string   `xml:"http://purl.org/dc/elements/1.1/ identifier,omitempty"`
	Keywords       string   `xml:"keywords,omitempty"`
	Language       string   `xml:"http://purl.org/dc/elements/1.1/ language,omitempty"`
	LastModifiedBy string   `xml:"lastModifiedBy,omitempty"`
	LastPrinted    *DCDate  `xml:"lastPrinted,omitempty"`
	Modified       *DCDate  `xml:"http://purl.org/dc/terms/ modified,omitempty"`
	Revision       string   `xml:"revision,omitempty"`
	Subject        string   `xml:"http://purl.org/dc/elements/1.1/ subject,omitempty"`
	Title          string   `xml:"http://purl.org/dc/elements/1.1/ title,omitempty"`
	Version        string   `xml:"version,omitempty"`
}

// DCDate represents a Dublin Core date with type attribute.
type DCDate struct {
	Type  string `xml:"http://www.w3.org/2001/XMLSchema-instance type,attr,omitempty"`
	Value string `xml:",chardata"`
}

// NewCoreProperties creates a new CoreProperties with defaults.
func NewCoreProperties() *CoreProperties {
	return &CoreProperties{}
}

// ExtendedProperties represents document extended properties.
type ExtendedProperties struct {
	XMLName          xml.Name `xml:"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties Properties"`
	Template         string   `xml:"Template,omitempty"`
	TotalTime        int      `xml:"TotalTime,omitempty"`
	Pages            int      `xml:"Pages,omitempty"`
	Words            int      `xml:"Words,omitempty"`
	Characters       int      `xml:"Characters,omitempty"`
	Application      string   `xml:"Application,omitempty"`
	DocSecurity      int      `xml:"DocSecurity,omitempty"`
	Lines            int      `xml:"Lines,omitempty"`
	Paragraphs       int      `xml:"Paragraphs,omitempty"`
	ScaleCrop        bool     `xml:"ScaleCrop,omitempty"`
	HeadingPairs     *HeadingPairs `xml:"HeadingPairs,omitempty"`
	TitlesOfParts    *TitlesOfParts `xml:"TitlesOfParts,omitempty"`
	Company          string   `xml:"Company,omitempty"`
	LinksUpToDate    bool     `xml:"LinksUpToDate,omitempty"`
	CharactersWithSpaces int  `xml:"CharactersWithSpaces,omitempty"`
	SharedDoc        bool     `xml:"SharedDoc,omitempty"`
	HyperlinksChanged bool    `xml:"HyperlinksChanged,omitempty"`
	AppVersion       string   `xml:"AppVersion,omitempty"`
}

// HeadingPairs represents heading pairs.
type HeadingPairs struct {
	Vector *Vector `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes vector"`
}

// TitlesOfParts represents titles of parts.
type TitlesOfParts struct {
	Vector *Vector `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes vector"`
}

// Vector represents a vector of values.
type Vector struct {
	Size     int            `xml:"size,attr"`
	BaseType string         `xml:"baseType,attr"`
	Variant  []*Variant     `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes variant,omitempty"`
	Lpstr    []string       `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes lpstr,omitempty"`
}

// Variant represents a variant value.
type Variant struct {
	Lpstr string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes lpstr,omitempty"`
	I4    int    `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes i4,omitempty"`
}
