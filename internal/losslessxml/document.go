// Package losslessxml inspects XML without reserializing it. It exposes bounded
// plain-text leaf edits; structural edits require separate ownership policies.
package losslessxml

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"sort"
	"strings"
	"unicode/utf8"
)

const xmlNS = "http://www.w3.org/XML/1998/namespace"
const xmlnsNS = "http://www.w3.org/2000/xmlns/"

type node struct {
	name                                 xml.Name
	qualified                            string
	attrs                                []xml.Attr
	attrNamespaces                       map[string]string
	start, contentStart, contentEnd, end int
	parent                               int
	leaf, selfClosing                    bool
	text                                 string
	ns                                   map[string]string
}
type Document struct {
	source []byte
	nodes  []node
}

// Limits are optional ceilings for an untrusted XML parse. Parse retains its
// historical defaults; callers of ParseWithLimits provide positive ceilings.
type Limits struct {
	MaxBytes       int
	MaxNodes       int
	MaxDepth       int
	MaxSourceUnits int // UTF-16 source units, before any parse allocation
}

// Refusal carries the parser's own failure category, not a classification of
// diagnostic message text. Cause may be inspected with errors.As.
type Refusal struct {
	Category string
	Cause    error
}

func (r *Refusal) Error() string { return r.Cause.Error() }
func (r *Refusal) Unwrap() error { return r.Cause }

// Element is an immutable handle valid only against its originating snapshot.
type Element struct {
	doc   *Document
	index int
}
type TextEdit struct {
	Target Element
	Text   string
}

// MismatchedTagError identifies a mismatched closing QName in a parsed source.
// Other parse failures retain their original errors and are not assigned this type.
type MismatchedTagError struct {
	Offset int64
}

func (e *MismatchedTagError) Error() string {
	return fmt.Sprintf("XML offset %d: end tag QName differs from start", e.Offset)
}

type frame struct {
	index int
	raw   xml.Name
	ns    map[string]string
	text  strings.Builder
}

func Parse(source []byte) (*Document, error) { return parse(source, Limits{}, false) }

// ParseWithLimits checks input size before copying source bytes and node/depth
// ceilings during scanning, before admitting the next element into the model.
func ParseWithLimits(source []byte, limits Limits) (*Document, error) {
	return parse(source, limits, true)
}

func parse(source []byte, limits Limits, bounded bool) (*Document, error) {
	refuse := func(category, reason string) (*Document, error) {
		return nil, &Refusal{Category: category, Cause: fmt.Errorf("%s", reason)}
	}
	if bounded && (limits.MaxNodes <= 0 || limits.MaxDepth <= 0 || limits.MaxSourceUnits <= 0 || limits.MaxBytes < 0) {
		return refuse("invalid-limit", "XML limits must be positive")
	}
	if limits.MaxBytes > 0 && len(source) > limits.MaxBytes {
		return refuse("input-too-large", "XML input length limit exceeded")
	}
	if !utf8.Valid(source) {
		return refuse("invalid-character", "XML input must be UTF-8")
	}
	if limits.MaxSourceUnits > 0 {
		units := 0
		remaining := source
		for len(remaining) > 0 {
			r, width := utf8.DecodeRune(remaining)
			if r > 0xffff {
				units += 2
			} else {
				units++
			}
			if units > limits.MaxSourceUnits {
				return refuse("input-too-large", "XML source-unit limit exceeded")
			}
			remaining = remaining[width:]
		}
	}
	depth := 512
	if bounded {
		depth = limits.MaxDepth
	}
	d := &Document{source: bytes.Clone(source)}
	decoderSource := d.source
	restoreName := func(n xml.Name) xml.Name { return n }
	if bounded {
		var err error
		decoderSource, restoreName, err = decoderNameCompatibility(d.source)
		if err != nil {
			return nil, err
		}
	}
	decoder := xml.NewDecoder(bytes.NewReader(decoderSource))
	var stack []*frame
	roots := 0
	fail := func(category, reason string) (*Document, error) {
		return refuse(category, fmt.Sprintf("XML offset %d: %s", decoder.InputOffset(), reason))
	}
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.RawToken()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			if bounded {
				return nil, classifyDecoderFailure(err, d.source[before:after])
			}
			return nil, err
		}
		switch t := token.(type) {
		case xml.StartElement:
			t.Name = restoreName(t.Name)
			for i := range t.Attr {
				t.Attr[i].Name = restoreName(t.Attr[i].Name)
			}
			if bounded && !lexicalAttributeSpacing(d.source[before:after]) {
				return fail("malformed-xml", "attribute whitespace missing")
			}
			t.Attr, err = normalizeAttributeValues(d.source[before:after], t.Attr)
			if err != nil {
				if bounded {
					return fail("malformed-xml", err.Error())
				}
				return nil, err
			}
			if len(stack) >= depth {
				return fail("depth-limit", "nesting limit exceeded")
			}
			if limits.MaxNodes > 0 && len(d.nodes) >= limits.MaxNodes {
				return fail("node-limit", "node limit exceeded")
			}
			ns := map[string]string{"xml": xmlNS}
			parent := -1
			if len(stack) > 0 {
				f := stack[len(stack)-1]
				parent = f.index
				d.nodes[parent].leaf = false
				for k, v := range f.ns {
					ns[k] = v
				}
			} else {
				roots++
				if roots > 1 {
					return fail("malformed-xml", "multiple roots")
				}
			}
			declared := map[string]bool{}
			for _, a := range t.Attr {
				prefix, isDecl := "", false
				if a.Name.Space == "xmlns" {
					prefix = a.Name.Local
					isDecl = true
				} else if a.Name.Space == "" && a.Name.Local == "xmlns" {
					isDecl = true
				}
				if !isDecl {
					continue
				}
				if prefix != "" && !localName(prefix) {
					return fail("malformed-xml", "invalid namespace prefix")
				}
				if declared[prefix] {
					return fail("duplicate-attribute", "duplicate namespace declaration")
				}
				declared[prefix] = true
				if prefix == "xmlns" || a.Value == xmlnsNS || (prefix == "xml" && a.Value != xmlNS) || (prefix != "xml" && a.Value == xmlNS) || (prefix != "" && a.Value == "") {
					return fail("malformed-xml", "invalid reserved namespace binding")
				}
				ns[prefix] = a.Value
			}
			name, err := expanded(t.Name, ns, false)
			if err != nil {
				category := "malformed-xml"
				if errors.Is(err, errUnboundPrefix) {
					category = "unbound-prefix"
				}
				return fail(category, err.Error())
			}
			attrs := []xml.Attr{}
			attrNamespaces := map[string]string{}
			seen := map[xml.Name]bool{}
			for _, a := range t.Attr {
				if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
					continue
				}
				resolved, err := expanded(a.Name, ns, true)
				if err != nil {
					category := "malformed-xml"
					if errors.Is(err, errUnboundPrefix) {
						category = "unbound-prefix"
					}
					return fail(category, err.Error())
				}
				if seen[resolved] {
					return fail("duplicate-attribute", "duplicate expanded attribute")
				}
				seen[resolved] = true
				attrs = append(attrs, xml.Attr{Name: resolved, Value: a.Value})
				qualified := a.Name.Local
				if a.Name.Space != "" {
					qualified = a.Name.Space + ":" + qualified
				}
				attrNamespaces[qualified] = resolved.Space
			}
			self := after-before >= 2 && bytes.HasSuffix(d.source[before:after], []byte("/>"))
			i := len(d.nodes)
			qualified := t.Name.Local
			if t.Name.Space != "" {
				qualified = t.Name.Space + ":" + qualified
			}
			d.nodes = append(d.nodes, node{name: name, qualified: qualified, attrs: attrs, attrNamespaces: attrNamespaces, start: before, contentStart: after, parent: parent, leaf: true, selfClosing: self, ns: ns})
			stack = append(stack, &frame{index: i, raw: t.Name, ns: ns})
		case xml.EndElement:
			t.Name = restoreName(t.Name)
			if len(stack) == 0 {
				return fail("malformed-xml", "end tag without start")
			}
			f := stack[len(stack)-1]
			if f.raw != t.Name {
				return nil, &MismatchedTagError{Offset: decoder.InputOffset()}
			}
			n := &d.nodes[f.index]
			n.contentEnd = before
			n.end = after
			n.text = f.text.String()
			if n.selfClosing {
				n.contentEnd = n.contentStart
				n.end = n.contentStart
			}
			stack = stack[:len(stack)-1]
		case xml.CharData:
			if len(stack) == 0 {
				// XML prolog/epilog allow literal S only. RawToken decodes
				// references and CDATA, which are not legal here even when
				// their decoded content is empty or whitespace.
				if !xmlWhitespace(d.source[before:after]) {
					return fail("malformed-xml", "non-literal whitespace outside root")
				}
			} else {
				stack[len(stack)-1].text.Write(t)
			}
		case xml.Comment:
			if len(stack) > 0 {
				d.nodes[stack[len(stack)-1].index].leaf = false
			}
		case xml.ProcInst:
			if strings.EqualFold(t.Target, "xml") {
				if t.Target != "xml" || before != 0 || roots != 0 {
					return fail("malformed-xml", "misplaced XML declaration")
				}
				if bounded && !validXMLDeclaration(t.Inst) {
					return fail("malformed-xml", "invalid XML declaration")
				}
			} else {
				if bounded && bytes.HasSuffix(d.source[before:after], []byte("/?>")) {
					return fail("malformed-xml", "invalid processing instruction")
				}
				if len(stack) > 0 {
					d.nodes[stack[len(stack)-1].index].leaf = false
				}
			}
		case xml.Directive:
			return fail("dtd-forbidden", "XML directives are unsupported")
		}
	}
	if roots != 1 || len(stack) != 0 {
		return fail("malformed-xml", "expected one complete root")
	}
	return d, nil
}

func expanded(raw xml.Name, ns map[string]string, attribute bool) (xml.Name, error) {
	if !localName(raw.Local) || raw.Space != "" && !localName(raw.Space) {
		return xml.Name{}, fmt.Errorf("invalid QName")
	}
	if raw.Space == "xmlns" {
		return xml.Name{}, fmt.Errorf("reserved xmlns prefix")
	}
	if raw.Space != "" {
		uri, ok := ns[raw.Space]
		if !ok || uri == "" {
			return xml.Name{}, unboundPrefix(raw.Space)
		}
		return xml.Name{Space: uri, Local: raw.Local}, nil
	}
	uri := ""
	if !attribute {
		uri = ns[""]
	}
	return xml.Name{Space: uri, Local: raw.Local}, nil
}
func (d *Document) Elements() []Element {
	out := make([]Element, len(d.nodes))
	for i := range out {
		out[i] = Element{d, i}
	}
	return out
}
func (e Element) valid() bool { return e.doc != nil && e.index >= 0 && e.index < len(e.doc.nodes) }

// Ordinal is source-order identity within a fingerprint-verified snapshot only.
func (e Element) Ordinal() int {
	if !e.valid() {
		return -1
	}
	return e.index
}

func (e Element) QualifiedName() string {
	if !e.valid() {
		return ""
	}
	return e.doc.nodes[e.index].qualified
}
func (e Element) Name() xml.Name {
	if !e.valid() {
		return xml.Name{}
	}
	return e.doc.nodes[e.index].name
}
func (e Element) Attributes() []xml.Attr {
	if !e.valid() {
		return nil
	}
	return append([]xml.Attr(nil), e.doc.nodes[e.index].attrs...)
}

// AttributeNamespaces returns a detached qualified-name to namespace-URI map.
// Mutation cannot change later reads from this parsed snapshot.
func (e Element) AttributeNamespaces() map[string]string {
	out := map[string]string{}
	if e.valid() {
		for k, v := range e.doc.nodes[e.index].attrNamespaces {
			out[k] = v
		}
	}
	return out
}
func (e Element) Parent() (Element, bool) {
	if !e.valid() {
		return Element{}, false
	}
	p := e.doc.nodes[e.index].parent
	if p < 0 {
		return Element{}, false
	}
	return Element{e.doc, p}, true
}

// SourceRange returns the original UTF-8 byte offsets [start, end) for this
// element's complete markup. The range is valid only for its source snapshot.
// An invalid handle returns (-1, -1).
func (e Element) SourceRange() (start, end int) {
	if !e.valid() {
		return -1, -1
	}
	n := e.doc.nodes[e.index]
	return n.start, n.end
}

// ContentRange is the original UTF-8 byte interval inside the element tags.
// For a self-closing element it is empty at the start of the closing slash.
func (e Element) ContentRange() (start, end int) {
	if !e.valid() {
		return -1, -1
	}
	n := e.doc.nodes[e.index]
	return n.contentStart, n.contentEnd
}

func (e Element) SelfClosing() bool {
	return e.valid() && e.doc.nodes[e.index].selfClosing
}

// Raw returns a private copy of the complete original element markup.
// It is evidence for conservative equality, never authority to splice elsewhere.
// Namespaces returns copied in-scope bindings for conservative QName equality.
func (e Element) Namespaces() map[string]string {
	out := map[string]string{}
	if !e.valid() {
		return out
	}
	for k, v := range e.doc.nodes[e.index].ns {
		out[k] = v
	}
	return out
}

func (e Element) Raw() []byte {
	if !e.valid() {
		return nil
	}
	n := e.doc.nodes[e.index]
	return bytes.Clone(e.doc.source[n.start:n.end])
}

func (e Element) Text() (string, bool) {
	if !e.valid() {
		return "", false
	}
	n := e.doc.nodes[e.index]
	return n.text, n.leaf
}

// ReplaceText returns a new document payload without mutating this snapshot.
// Leaf content may contain entity references or CDATA; changed values are escaped
// XML character data. No-ops retain original entity/CDATA spellings. Comments,
// processing instructions, child elements and self-closing insertions refuse.
func (d *Document) ReplaceText(edits []TextEdit) ([]byte, error) {
	type splice struct {
		start, end int
		data       []byte
	}
	var changes []splice
	seen := map[int]bool{}
	for _, edit := range edits {
		e := edit.Target
		if e.doc != d || !e.valid() {
			return nil, fmt.Errorf("foreign or invalid target")
		}
		if seen[e.index] {
			return nil, fmt.Errorf("duplicate target")
		}
		seen[e.index] = true
		n := d.nodes[e.index]
		if !n.leaf {
			return nil, fmt.Errorf("target is not a plain text leaf")
		}
		if !validText(edit.Text) {
			return nil, fmt.Errorf("invalid XML character in replacement")
		}
		if edit.Text == n.text {
			continue
		}
		if n.selfClosing {
			return nil, fmt.Errorf("self-closing insertion requires structural edit")
		}
		var b bytes.Buffer
		if err := xml.EscapeText(&b, []byte(edit.Text)); err != nil {
			return nil, err
		}
		changes = append(changes, splice{n.contentStart, n.contentEnd, b.Bytes()})
	}
	sort.Slice(changes, func(i, j int) bool { return changes[i].start < changes[j].start })
	var out bytes.Buffer
	cursor := 0
	for _, s := range changes {
		if s.start < cursor {
			return nil, fmt.Errorf("overlapping edits")
		}
		out.Write(d.source[cursor:s.start])
		out.Write(s.data)
		cursor = s.end
	}
	out.Write(d.source[cursor:])
	return out.Bytes(), nil
}
func validText(s string) bool {
	if !utf8.ValidString(s) {
		return false
	}
	for _, r := range s {
		if r == 9 || r == 10 || r == 13 || r >= 0x20 && r <= 0xd7ff || r >= 0xe000 && r <= 0xfffd || r >= 0x10000 && r <= 0x10ffff {
			continue
		}
		return false
	}
	return true
}

func xmlWhitespace(data []byte) bool {
	for _, b := range data {
		if b != ' ' && b != '\t' && b != '\r' && b != '\n' {
			return false
		}
	}
	return true
}
