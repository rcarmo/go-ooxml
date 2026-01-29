package packaging

import (
	"encoding/xml"
	"path"
	"strings"
)

// ContentTypes represents the [Content_Types].xml file.
type ContentTypes struct {
	XMLName   xml.Name   `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Defaults  []Default  `xml:"Default"`
	Overrides []Override `xml:"Override"`
}

// Default represents a default content type mapping by extension.
type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// Override represents an override content type mapping by part name.
type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// NewContentTypes creates a new ContentTypes with standard defaults.
func NewContentTypes() *ContentTypes {
	return &ContentTypes{
		Defaults: []Default{
			{Extension: "rels", ContentType: ContentTypeRelationships},
			{Extension: "xml", ContentType: ContentTypeXML},
			{Extension: "png", ContentType: ContentTypePNG},
			{Extension: "jpeg", ContentType: ContentTypeJPEG},
			{Extension: "jpg", ContentType: ContentTypeJPEG},
			{Extension: "gif", ContentType: ContentTypeGIF},
		},
		Overrides: make([]Override, 0),
	}
}

// GetContentType returns the content type for a given part URI.
func (ct *ContentTypes) GetContentType(uri string) string {
	// Normalize URI to start with /
	if !strings.HasPrefix(uri, "/") {
		uri = "/" + uri
	}

	// Check overrides first (exact match)
	for _, o := range ct.Overrides {
		if o.PartName == uri {
			return o.ContentType
		}
	}

	// Fall back to defaults by extension
	ext := strings.TrimPrefix(path.Ext(uri), ".")
	for _, d := range ct.Defaults {
		if strings.EqualFold(d.Extension, ext) {
			return d.ContentType
		}
	}

	return ""
}

// AddOverride adds a content type override for a specific part.
func (ct *ContentTypes) AddOverride(partName, contentType string) {
	// Normalize to start with /
	if !strings.HasPrefix(partName, "/") {
		partName = "/" + partName
	}

	// Check if already exists
	for i, o := range ct.Overrides {
		if o.PartName == partName {
			ct.Overrides[i].ContentType = contentType
			return
		}
	}

	ct.Overrides = append(ct.Overrides, Override{
		PartName:    partName,
		ContentType: contentType,
	})
}

// AddDefault adds a default content type for an extension.
func (ct *ContentTypes) AddDefault(extension, contentType string) {
	// Remove leading dot if present
	extension = strings.TrimPrefix(extension, ".")

	// Check if already exists
	for i, d := range ct.Defaults {
		if strings.EqualFold(d.Extension, extension) {
			ct.Defaults[i].ContentType = contentType
			return
		}
	}

	ct.Defaults = append(ct.Defaults, Default{
		Extension:   extension,
		ContentType: contentType,
	})
}

// RemoveOverride removes a content type override.
func (ct *ContentTypes) RemoveOverride(partName string) bool {
	if !strings.HasPrefix(partName, "/") {
		partName = "/" + partName
	}

	for i, o := range ct.Overrides {
		if o.PartName == partName {
			ct.Overrides = append(ct.Overrides[:i], ct.Overrides[i+1:]...)
			return true
		}
	}
	return false
}

// EnsureContentType ensures a content type is set for a part,
// adding an override if the extension default doesn't match.
func (ct *ContentTypes) EnsureContentType(partName, contentType string) {
	current := ct.GetContentType(partName)
	if current != contentType {
		ct.AddOverride(partName, contentType)
	}
}
