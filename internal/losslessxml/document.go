// Package losslessxml inspects XML without reserializing it. It exposes bounded
// plain-text leaf edits; structural edits require separate ownership policies.
package losslessxml

import (
	"bytes"
	"encoding/xml"
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
	attrs                                []xml.Attr
	start, contentStart, contentEnd, end int
	parent                               int
	leaf, selfClosing                    bool
	text                                 string
}
type Document struct {
	source []byte
	nodes  []node
}

// Element is an immutable handle valid only against its originating snapshot.
type Element struct {
	doc   *Document
	index int
}
type TextEdit struct {
	Target Element
	Text   string
}
type frame struct {
	index int
	raw   xml.Name
	ns    map[string]string
	text  strings.Builder
}

func Parse(source []byte) (*Document, error) {
	if !utf8.Valid(source) {
		return nil, fmt.Errorf("XML input must be UTF-8")
	}
	d := &Document{source: bytes.Clone(source)}
	decoder := xml.NewDecoder(bytes.NewReader(d.source))
	var stack []*frame
	roots := 0
	fail := func(reason string) (*Document, error) {
		return nil, fmt.Errorf("XML offset %d: %s", decoder.InputOffset(), reason)
	}
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.RawToken()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := token.(type) {
		case xml.StartElement:
			if len(stack) >= 512 {
				return fail("nesting limit exceeded")
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
					return fail("multiple roots")
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
				if declared[prefix] {
					return fail("duplicate namespace declaration")
				}
				declared[prefix] = true
				if prefix == "xmlns" || a.Value == xmlnsNS || (prefix == "xml" && a.Value != xmlNS) || (prefix != "xml" && a.Value == xmlNS) || (prefix != "" && a.Value == "") {
					return fail("invalid reserved namespace binding")
				}
				ns[prefix] = a.Value
			}
			name, err := expanded(t.Name, ns, false)
			if err != nil {
				return fail(err.Error())
			}
			attrs := []xml.Attr{}
			seen := map[xml.Name]bool{}
			for _, a := range t.Attr {
				if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
					continue
				}
				resolved, err := expanded(a.Name, ns, true)
				if err != nil {
					return fail(err.Error())
				}
				if seen[resolved] {
					return fail("duplicate expanded attribute")
				}
				seen[resolved] = true
				attrs = append(attrs, xml.Attr{Name: resolved, Value: a.Value})
			}
			self := after-before >= 2 && bytes.HasSuffix(d.source[before:after], []byte("/>"))
			i := len(d.nodes)
			d.nodes = append(d.nodes, node{name: name, attrs: attrs, start: before, contentStart: after, parent: parent, leaf: true, selfClosing: self})
			stack = append(stack, &frame{index: i, raw: t.Name, ns: ns})
		case xml.EndElement:
			if len(stack) == 0 {
				return fail("end tag without start")
			}
			f := stack[len(stack)-1]
			if f.raw != t.Name {
				return fail("end tag QName differs from start")
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
				if len(bytes.TrimSpace(t)) != 0 {
					return fail("text outside root")
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
					return fail("misplaced XML declaration")
				}
				// Let the ordinary decoder enforce XML version/encoding declarations too.
			} else if len(stack) > 0 {
				d.nodes[stack[len(stack)-1].index].leaf = false
			}
		case xml.Directive:
			return fail("XML directives are unsupported")
		}
	}
	if roots != 1 || len(stack) != 0 {
		return fail("expected one complete root")
	}
	return d, nil
}

func expanded(raw xml.Name, ns map[string]string, attribute bool) (xml.Name, error) {
	if strings.Contains(raw.Local, ":") || strings.Contains(raw.Space, ":") || raw.Local == "" {
		return xml.Name{}, fmt.Errorf("invalid QName")
	}
	if raw.Space == "xmlns" {
		return xml.Name{}, fmt.Errorf("reserved xmlns prefix")
	}
	if raw.Space != "" {
		uri, ok := ns[raw.Space]
		if !ok || uri == "" {
			return xml.Name{}, fmt.Errorf("undeclared prefix %s", raw.Space)
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
