// Package xmlsnapshot exposes read-only XML snapshots without rewriting source bytes.
// Offsets and edit operations remain in internal/losslessxml; this package does not
// define a public source-offset contract.
package xmlsnapshot

import (
	"bytes"
	"encoding/xml"
	"errors"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const CategoryMalformedXML = "malformed-xml"

// ParseError is returned only for natively classified failures. Other parser
// errors are returned as-is, not indiscriminately labelled malformed XML.
type ParseError struct {
	Category string
	Cause    error
}

func (e *ParseError) Error() string { return e.Cause.Error() }
func (e *ParseError) Unwrap() error { return e.Cause }

// Document owns an immutable snapshot of the parsed source.
type Document struct {
	source []byte
	doc    *losslessxml.Document
}

// Element is a read-only view of an element in a Document.
type Element struct {
	doc     *losslessxml.Document
	element losslessxml.Element
}

// Parse copies caller bytes and returns no document on any failure.
func Parse(source []byte) (*Document, error) {
	parsed, err := losslessxml.Parse(source)
	if err != nil {
		var mismatch *losslessxml.MismatchedTagError
		var syntax *xml.SyntaxError
		if errors.As(err, &mismatch) || errors.As(err, &syntax) {
			return nil, &ParseError{Category: CategoryMalformedXML, Cause: err}
		}
		return nil, err
	}
	return &Document{source: bytes.Clone(source), doc: parsed}, nil
}

// Source returns a private copy; modifying it cannot change the parsed model.
func (d *Document) Source() []byte {
	if d == nil {
		return nil
	}
	return bytes.Clone(d.source)
}

// Root returns the sole parsed root, or an empty view for a nil document.
func (d *Document) Root() Element {
	if d == nil || d.doc == nil {
		return Element{}
	}
	elements := d.doc.Elements()
	if len(elements) == 0 {
		return Element{}
	}
	return Element{doc: d.doc, element: elements[0]}
}

func (e Element) Name() xml.Name       { return e.element.Name() }
func (e Element) Text() (string, bool) { return e.element.Text() }

// Children returns direct children in source order without exposing mutable nodes.
func (e Element) Children() []Element {
	if e.doc == nil || e.element.Ordinal() < 0 {
		return nil
	}
	var children []Element
	for _, child := range e.doc.Elements() {
		parent, ok := child.Parent()
		if ok && parent.Ordinal() == e.element.Ordinal() {
			children = append(children, Element{doc: e.doc, element: child})
		}
	}
	return children
}

// Attributes returns a detached slice of decoded expanded XML attributes.
func (e Element) Attributes() []xml.Attr { return e.element.Attributes() }

// Attribute returns a decoded value by expanded name; absence is explicit.
func (e Element) Attribute(namespace, local string) (string, bool) {
	for _, attr := range e.element.Attributes() {
		if attr.Name == (xml.Name{Space: namespace, Local: local}) {
			return attr.Value, true
		}
	}
	return "", false
}

// AttributeNamespaces returns detached qualified-attribute namespace metadata.
// Mutating or deleting entries affects only the returned map, not this model.
func (e Element) AttributeNamespaces() map[string]string {
	return e.element.AttributeNamespaces()
}
