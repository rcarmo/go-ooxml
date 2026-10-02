package losslessxml

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
	"unicode/utf8"
)

// UnboundAttributeNamespaceError preserves the legacy diagnostic while allowing
// uniform callers to classify an authored namespace refusal without message text.
type UnboundAttributeNamespaceError struct{ Cause error }

func (e *UnboundAttributeNamespaceError) Error() string { return e.Cause.Error() }
func (e *UnboundAttributeNamespaceError) Unwrap() error { return e.Cause }

// AttributeEdit changes or adds one non-namespace attribute by expanded name.
// A new prefixed attribute requires an already in-scope namespace binding.
type AttributeEdit struct {
	Target Element
	Name   xml.Name
	Value  string
}
type attrRange struct {
	name       xml.Name
	start, end int
	quote      byte
}
type byteSplice struct {
	start, end int
	data       []byte
}

// Edit applies text and attribute updates atomically against one immutable source.
// Namespace declarations cannot be changed. Existing attribute quotes/spacing and
// every unrelated byte are retained; insertions use an already-bound prefix.
func (d *Document) Edit(texts []TextEdit, attributes []AttributeEdit) ([]byte, error) {
	// Validate text rules once. Edits below are collected using original offsets.
	if _, err := d.ReplaceText(texts); err != nil {
		return nil, err
	}
	splices := []byteSplice{}
	for _, t := range texts {
		n := d.nodes[t.Target.index]
		if t.Text == n.text {
			continue
		}
		var b bytes.Buffer
		_ = xml.EscapeText(&b, []byte(t.Text))
		splices = append(splices, byteSplice{n.contentStart, n.contentEnd, b.Bytes()})
	}
	type identity struct {
		node int
		name xml.Name
	}
	seen := map[identity]bool{}
	insertions := map[int][]byte{}
	for _, a := range attributes {
		if a.Target.doc != d || !a.Target.valid() {
			return nil, fmt.Errorf("foreign attribute target")
		}
		n := d.nodes[a.Target.index]
		if !localName(a.Name.Local) || a.Name.Local == "xmlns" || a.Name.Space == xmlnsNS {
			return nil, fmt.Errorf("invalid attribute name or namespace mutation")
		}
		if !validText(a.Value) {
			return nil, fmt.Errorf("invalid XML attribute text")
		}
		key := identity{a.Target.index, a.Name}
		if seen[key] {
			return nil, fmt.Errorf("duplicate attribute edit")
		}
		seen[key] = true
		existing := false
		same := false
		for _, v := range n.attrs {
			if v.Name == a.Name {
				existing = true
				same = v.Value == a.Value
				break
			}
		}
		if same {
			continue
		}
		ranges, err := d.attributeRanges(n)
		if err != nil {
			return nil, err
		}
		var b bytes.Buffer
		_ = xml.EscapeText(&b, []byte(a.Value))
		if existing {
			found := false
			for _, r := range ranges {
				if r.name == a.Name {
					splices = append(splices, byteSplice{r.start, r.end, b.Bytes()})
					found = true
					break
				}
			}
			if !found {
				return nil, fmt.Errorf("attribute lexical mapping missing")
			}
			continue
		}
		qname := a.Name.Local
		if a.Name.Space != "" {
			prefixes := []string{}
			for prefix, uri := range n.ns {
				if prefix != "" && uri == a.Name.Space {
					prefixes = append(prefixes, prefix)
				}
			}
			sort.Strings(prefixes)
			if len(prefixes) == 0 {
				return nil, &UnboundAttributeNamespaceError{Cause: fmt.Errorf("attribute namespace has no in-scope prefix")}
			}
			qname = prefixes[0] + ":" + qname
		}
		at := n.contentStart - 1
		if n.selfClosing {
			at--
		}
		insertions[at] = append(insertions[at], []byte(" "+qname+"=\""+b.String()+"\"")...)
	}
	for at, data := range insertions {
		splices = append(splices, byteSplice{at, at, data})
	}
	sort.SliceStable(splices, func(i, j int) bool { return splices[i].start < splices[j].start })
	var out bytes.Buffer
	cursor := 0
	for _, s := range splices {
		if s.start < cursor {
			return nil, fmt.Errorf("overlapping XML edits")
		}
		out.Write(d.source[cursor:s.start])
		out.Write(s.data)
		cursor = s.end
	}
	out.Write(d.source[cursor:])
	if _, err := Parse(out.Bytes()); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// ValidAuthoredName checks the existing structured edit name and value grammar.
// Exporting the validator lets bounded profiles classify refusal before rendering.
func ValidAuthoredName(name xml.Name, attribute bool) bool {
	return localName(name.Local) && name.Space != xmlnsNS && (!attribute || name.Local != "xmlns") && validText(name.Space)
}

func ValidAuthoredText(s string) bool { return validText(s) }

// XML 1.0 fifth-edition NCName productions (Name without colon).
func localName(s string) bool {
	if s == "" || !utf8.ValidString(s) {
		return false
	}
	for i, r := range s {
		if nameStart(r) {
			continue
		}
		if i == 0 || !(r == '-' || r == '.' || r >= '0' && r <= '9' || r == 0xb7 || r >= 0x300 && r <= 0x36f || r >= 0x203f && r <= 0x2040) {
			return false
		}
	}
	return true
}
func nameStart(r rune) bool {
	return r == '_' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' || r >= 0xc0 && r <= 0xd6 || r >= 0xd8 && r <= 0xf6 || r >= 0xf8 && r <= 0x2ff || r >= 0x370 && r <= 0x37d || r >= 0x37f && r <= 0x1fff || r >= 0x200c && r <= 0x200d || r >= 0x2070 && r <= 0x218f || r >= 0x2c00 && r <= 0x2fef || r >= 0x3001 && r <= 0xd7ff || r >= 0xf900 && r <= 0xfdcf || r >= 0xfdf0 && r <= 0xfffd || r >= 0x10000 && r <= 0xeffff
}
func (d *Document) attributeRanges(n node) ([]attrRange, error) {
	src := d.source
	at := n.start + 1
	stop := n.contentStart
	space := func(b byte) bool { return b == ' ' || b == '\t' || b == '\r' || b == '\n' }
	for at < stop && !space(src[at]) && src[at] != '>' && src[at] != '/' {
		at++
	}
	out := []attrRange{}
	for at < stop {
		for at < stop && space(src[at]) {
			at++
		}
		if at >= stop || src[at] == '>' || src[at] == '/' {
			break
		}
		start := at
		for at < stop && !space(src[at]) && src[at] != '=' {
			at++
		}
		qname := string(src[start:at])
		for at < stop && space(src[at]) {
			at++
		}
		if at >= stop || src[at] != '=' {
			return nil, fmt.Errorf("attribute equals expected")
		}
		at++
		for at < stop && space(src[at]) {
			at++
		}
		if at >= stop || (src[at] != '\'' && src[at] != '"') {
			return nil, fmt.Errorf("attribute quote expected")
		}
		quote := src[at]
		at++
		valueStart := at
		for at < stop && src[at] != quote {
			at++
		}
		if at >= stop {
			return nil, fmt.Errorf("attribute quote unterminated")
		}
		valueEnd := at
		at++
		if qname == "xmlns" || strings.HasPrefix(qname, "xmlns:") {
			continue
		}
		prefix, local := "", qname
		if i := strings.IndexByte(qname, ':'); i >= 0 {
			prefix, local = qname[:i], qname[i+1:]
		}
		resolved, err := expanded(xml.Name{Space: prefix, Local: local}, n.ns, true)
		if err != nil {
			return nil, err
		}
		out = append(out, attrRange{resolved, valueStart, valueEnd, quote})
	}
	return out, nil
}
