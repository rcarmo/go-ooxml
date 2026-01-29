package utils

import (
	"bytes"
	"encoding/xml"
	"io"
	"strings"
)

// XMLHeader is the standard XML declaration.
const XMLHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

// MarshalXMLWithHeader marshals v to XML with the standard XML declaration.
func MarshalXMLWithHeader(v interface{}) ([]byte, error) {
	data, err := xml.Marshal(v)
	if err != nil {
		return nil, err
	}
	return append([]byte(XMLHeader), data...), nil
}

// MarshalXMLIndentWithHeader marshals v to indented XML with the standard XML declaration.
func MarshalXMLIndentWithHeader(v interface{}, prefix, indent string) ([]byte, error) {
	data, err := xml.MarshalIndent(v, prefix, indent)
	if err != nil {
		return nil, err
	}
	return append([]byte(XMLHeader), data...), nil
}

// UnmarshalXML unmarshals XML data, stripping BOM if present.
func UnmarshalXML(data []byte, v interface{}) error {
	// Strip UTF-8 BOM if present
	data = bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF})
	return xml.Unmarshal(data, v)
}

// NewXMLDecoder creates an XML decoder that handles common OOXML quirks.
func NewXMLDecoder(r io.Reader) *xml.Decoder {
	d := xml.NewDecoder(r)
	d.CharsetReader = func(charset string, input io.Reader) (io.Reader, error) {
		// OOXML files are always UTF-8
		return input, nil
	}
	return d
}

// EscapeXMLText escapes special characters in XML text content.
func EscapeXMLText(s string) string {
	var buf strings.Builder
	xml.EscapeText(&buf, []byte(s))
	return buf.String()
}

// BoolPtr returns a pointer to a bool value.
func BoolPtr(v bool) *bool {
	return &v
}

// IntPtr returns a pointer to an int value.
func IntPtr(v int) *int {
	return &v
}

// Int64Ptr returns a pointer to an int64 value.
func Int64Ptr(v int64) *int64 {
	return &v
}

// StringPtr returns a pointer to a string value.
func StringPtr(v string) *string {
	return &v
}

// Float64Ptr returns a pointer to a float64 value.
func Float64Ptr(v float64) *float64 {
	return &v
}

// DerefBool returns the value of a bool pointer, or the default if nil.
func DerefBool(p *bool, def bool) bool {
	if p == nil {
		return def
	}
	return *p
}

// DerefString returns the value of a string pointer, or the default if nil.
func DerefString(p *string, def string) string {
	if p == nil {
		return def
	}
	return *p
}

// DerefInt returns the value of an int pointer, or the default if nil.
func DerefInt(p *int, def int) int {
	if p == nil {
		return def
	}
	return *p
}
