// Package theme provides DrawingML theme types.
package theme

import "encoding/xml"

// Namespaces used in DrawingML theme parts.
const (
	NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
)

// Theme represents a DrawingML theme part.
type Theme struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/drawingml/2006/main theme"`
	Name          string         `xml:"name,attr,omitempty"`
	ThemeElements *ThemeElements `xml:"themeElements,omitempty"`
}

// ThemeElements represents the theme elements collection.
type ThemeElements struct {
	ClrScheme *ClrScheme `xml:"clrScheme,omitempty"`
}

// ClrScheme represents a color scheme.
type ClrScheme struct {
	Name string `xml:"name,attr,omitempty"`
}
