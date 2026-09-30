// Package xmlsnapshot exposes immutable XML snapshots and lexical edits without
// rewriting untouched source bytes. Offsets address the original UTF-8 source.
package xmlsnapshot

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const (
	CategoryMalformedXML       = "malformed-xml"
	CategoryInvalidLimit       = "invalid-limit"
	CategoryInputTooLarge      = "input-too-large"
	CategoryDepthLimit         = "depth-limit"
	CategoryNodeLimit          = "node-limit"
	CategoryDTDForbidden       = "dtd-forbidden"
	CategoryEntityForbidden    = "entity-forbidden"
	CategoryDuplicateAttribute = "duplicate-attribute"
	CategoryUnboundPrefix      = "unbound-prefix"
	CategoryMismatchedTag      = "mismatched-tag"
	CategoryInvalidCharacter   = "invalid-character"
)

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
	source  []byte
	offsets []int // UTF-16 offset at rune boundaries; -1 within UTF-8 rune
	doc     *losslessxml.Document
}

// Limits is a fully specified bounded admission profile. All three XML
// ceilings must be positive. MaxBytes is an optional additional UTF-8 ceiling.
type Limits = losslessxml.Limits

// DefaultLimits is the native bounded-admission profile. The source ceiling
// counts UTF-16 code units, not UTF-8 bytes or Unicode scalar values.
func DefaultLimits() Limits {
	return Limits{MaxDepth: 256, MaxNodes: 100000, MaxSourceUnits: 8388608}
}

// LexicalLimits distinguishes omitted overrides from explicitly invalid zero.
// Nil fields use DefaultLimits; present values must be positive.
type LexicalLimits struct {
	MaxDepth, MaxNodes, MaxSourceUnits *int
}

// ParseLexical admits untrusted XML with finite defaults. Explicit zero or
// negative overrides refuse before any XML tree is returned.
func ParseLexical(source []byte, overrides LexicalLimits) (*Document, error) {
	limits := DefaultLimits()
	if overrides.MaxDepth != nil {
		limits.MaxDepth = *overrides.MaxDepth
	}
	if overrides.MaxNodes != nil {
		limits.MaxNodes = *overrides.MaxNodes
	}
	if overrides.MaxSourceUnits != nil {
		limits.MaxSourceUnits = *overrides.MaxSourceUnits
	}
	return ParseWithLimits(source, limits)
}

// Element is a read-only view of an element in a Document.
type Element struct {
	doc     *losslessxml.Document
	source  []byte
	offsets []int
	element losslessxml.Element
}

// Parse retains the original legacy admission behavior for existing callers.
// New lexical-profile callers use ParseLexical for finite default limits.
func Parse(source []byte) (*Document, error) {
	parsed, err := losslessxml.Parse(source)
	return wrapParsed(source, parsed, err, false)
}

// ParseWithLimits returns no document on refusal. Limits apply before a model
// is returned; the source and every returned value remain detached from input.
func ParseWithLimits(source []byte, limits Limits) (*Document, error) {
	parsed, err := losslessxml.ParseWithLimits(source, limits)
	return wrapParsed(source, parsed, err, true)
}

func wrapParsed(source []byte, parsed *losslessxml.Document, err error, typed bool) (*Document, error) {
	if err != nil {
		var refusal *losslessxml.Refusal
		var mismatch *losslessxml.MismatchedTagError
		var syntax *xml.SyntaxError
		switch {
		case typed && errors.As(err, &refusal):
			return nil, &ParseError{Category: refusal.Category, Cause: err}
		case errors.As(err, &mismatch):
			category := CategoryMalformedXML
			if typed {
				category = CategoryMismatchedTag
			}
			return nil, &ParseError{Category: category, Cause: err}
		case errors.As(err, &syntax):
			return nil, &ParseError{Category: CategoryMalformedXML, Cause: err}
		}
		return nil, err
	}
	copyOfSource := bytes.Clone(source)
	offsets := make([]int, len(copyOfSource)+1)
	for i := range offsets {
		offsets[i] = -1
	}
	units := 0
	for i := 0; i < len(copyOfSource); {
		offsets[i] = units
		r, width := utf8.DecodeRune(copyOfSource[i:])
		if r > 0xffff {
			units += 2
		} else {
			units++
		}
		i += width
	}
	offsets[len(copyOfSource)] = units
	return &Document{source: copyOfSource, offsets: offsets, doc: parsed}, nil
}

// Source returns a private copy; modifying it cannot change the parsed model.
// Elements returns source-order elements from the same immutable snapshot.
func (d *Document) Elements() []Element {
	if d == nil || d.doc == nil {
		return nil
	}
	all := d.doc.Elements()
	out := make([]Element, len(all))
	for i, e := range all {
		out[i] = Element{doc: d.doc, source: d.source, offsets: d.offsets, element: e}
	}
	return out
}

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
	elements := d.Elements()
	if len(elements) == 0 {
		return Element{}
	}
	return elements[0]
}

// SameElement compares snapshot identity and source-order element identity.
func (e Element) SameElement(other Element) bool {
	return e.doc != nil && e.doc == other.doc && e.element.Ordinal() >= 0 && e.element.Ordinal() == other.element.Ordinal()
}

func (e Element) Name() xml.Name        { return e.element.Name() }
func (e Element) QualifiedName() string { return e.element.QualifiedName() }
func (e Element) Text() (string, bool)  { return e.element.Text() }
func (e Element) SelfClosing() bool     { return e.element.SelfClosing() }

// OffsetRange is a half-open original-source interval, in both UTF-8 bytes and
// UTF-16 code units. It never indexes decoded character data.
type OffsetRange struct {
	ByteStart, ByteEnd   int
	UTF16Start, UTF16End int
}

func (d *Document) sourceRange(start, end int) (OffsetRange, error) {
	if d == nil || start < 0 || end < start || end > len(d.source) {
		return OffsetRange{}, fmt.Errorf("invalid XML source range")
	}
	if d.offsets[start] < 0 || d.offsets[end] < 0 {
		return OffsetRange{}, fmt.Errorf("XML source range splits UTF-8 scalar")
	}
	return OffsetRange{start, end, d.offsets[start], d.offsets[end]}, nil
}

// SourceRange addresses complete markup (including tags) in the original input.
func (e Element) SourceRange() (OffsetRange, error) {
	if e.doc == nil {
		return OffsetRange{}, fmt.Errorf("invalid XML element")
	}
	start, end := e.element.SourceRange()
	return (&Document{source: e.source, offsets: e.offsets, doc: e.doc}).sourceRange(start, end)
}

// ContentRange addresses the original bytes between start and end tags.
func (e Element) ContentRange() (OffsetRange, error) {
	if e.doc == nil {
		return OffsetRange{}, fmt.Errorf("invalid XML element")
	}
	start, end := e.element.ContentRange()
	return (&Document{source: e.source, offsets: e.offsets, doc: e.doc}).sourceRange(start, end)
}

// Parent returns the owning element, or false for the root/invalid view.
func (e Element) Root() Element {
	for parent, ok := e.Parent(); ok; parent, ok = e.Parent() {
		e = parent
	}
	return e
}

func (e Element) Parent() (Element, bool) {
	parent, ok := e.element.Parent()
	if !ok {
		return Element{}, false
	}
	return Element{doc: e.doc, source: e.source, offsets: e.offsets, element: parent}, true
}

// Children returns direct children in source order without exposing mutable nodes.
func (e Element) Children() []Element {
	if e.doc == nil || e.element.Ordinal() < 0 {
		return nil
	}
	var children []Element
	for _, child := range e.doc.Elements() {
		parent, ok := child.Parent()
		if ok && parent.Ordinal() == e.element.Ordinal() {
			children = append(children, Element{doc: e.doc, source: e.source, offsets: e.offsets, element: child})
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
