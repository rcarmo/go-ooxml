package losslessxml

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"sort"
)

// NewElement uses expanded names. Text and Children are mutually exclusive;
// namespace declarations are authored by the renderer, not supplied as attributes.
type NewElement struct {
	Name       xml.Name
	Attributes []xml.Attr
	Text       string
	Children   []NewElement
}
type ChildInsertion struct {
	Parent   Element
	Children []NewElement
}

// InsertChildren appends fully structured children atomically. Original bytes are
// preserved except a self-closing parent's slash is replaced by children/end tag.
// Duplicate or nested targets refuse. Existing bindings are never overwritten;
// newly allocated prefixes are scoped to inserted nodes. Raw XML is not accepted.
func (d *Document) InsertChildren(edits []ChildInsertion) ([]byte, error) {
	seen := map[int]bool{}
	var parents []Element
	for _, edit := range edits {
		e := edit.Parent
		if e.doc != d || !e.valid() {
			return nil, fmt.Errorf("foreign insertion parent")
		}
		if len(edit.Children) == 0 {
			continue
		}
		if seen[e.index] {
			return nil, fmt.Errorf("duplicate insertion parent")
		}
		seen[e.index] = true
		parents = append(parents, e)
	}
	for _, e := range parents {
		for p, ok := e.Parent(); ok; p, ok = p.Parent() {
			if seen[p.index] {
				return nil, fmt.Errorf("nested insertion targets")
			}
		}
	}
	splices := []byteSplice{}
	for _, edit := range edits {
		if len(edit.Children) == 0 {
			continue
		}
		n := d.nodes[edit.Parent.index]
		var inserted bytes.Buffer
		for _, child := range edit.Children {
			if err := renderElement(&inserted, child, n.ns, 0); err != nil {
				return nil, err
			}
		}
		if !n.selfClosing {
			splices = append(splices, byteSplice{n.contentEnd, n.contentEnd, inserted.Bytes()})
			continue
		}
		end := n.start + 1
		for end < n.contentStart {
			b := d.source[end]
			if b == ' ' || b == '\t' || b == '\r' || b == '\n' || b == '/' || b == '>' {
				break
			}
			end++
		}
		qname := d.source[n.start+1 : end]
		var replacement bytes.Buffer
		replacement.WriteByte('>')
		replacement.Write(inserted.Bytes())
		replacement.WriteString("</")
		replacement.Write(qname)
		replacement.WriteByte('>')
		splices = append(splices, byteSplice{n.contentStart - 2, n.contentStart, replacement.Bytes()})
	}
	sort.Slice(splices, func(i, j int) bool { return splices[i].start < splices[j].start })
	var out bytes.Buffer
	at := 0
	for _, s := range splices {
		if s.start < at {
			return nil, fmt.Errorf("overlapping insertion")
		}
		out.Write(d.source[at:s.start])
		out.Write(s.data)
		at = s.end
	}
	out.Write(d.source[at:])
	if _, err := Parse(out.Bytes()); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}
func renderElement(out *bytes.Buffer, node NewElement, inherited map[string]string, depth int) error {
	if depth >= 512 {
		return fmt.Errorf("inserted XML exceeds nesting limit")
	}
	if !localName(node.Name.Local) || node.Name.Space == xmlnsNS {
		return fmt.Errorf("invalid inserted element name")
	}
	if !validText(node.Text) || !validText(node.Name.Space) {
		return fmt.Errorf("invalid inserted text/namespace")
	}
	if node.Text != "" && len(node.Children) > 0 {
		return fmt.Errorf("mixed inserted content must have explicit structure")
	}
	ns := map[string]string{}
	for k, v := range inherited {
		ns[k] = v
	}
	type declaration struct{ prefix, uri string }
	var declarations []declaration
	qualify := func(name xml.Name, attribute bool) (string, error) {
		if !localName(name.Local) || name.Space == xmlnsNS || attribute && name.Local == "xmlns" {
			return "", fmt.Errorf("invalid name or namespace declaration")
		}
		if !validText(name.Space) {
			return "", fmt.Errorf("invalid namespace value")
		}
		if name.Space == "" {
			if !attribute && ns[""] != "" {
				ns[""] = ""
				declarations = append(declarations, declaration{"", ""})
			}
			return name.Local, nil
		}
		if !attribute && ns[""] == name.Space {
			return name.Local, nil
		}
		prefixes := []string{}
		for prefix, uri := range ns {
			if prefix != "" && uri == name.Space {
				prefixes = append(prefixes, prefix)
			}
		}
		sort.Strings(prefixes)
		if len(prefixes) > 0 {
			return prefixes[0] + ":" + name.Local, nil
		}
		prefix := ""
		for i := 1; ; i++ {
			candidate := fmt.Sprintf("n%d", i)
			if _, exists := ns[candidate]; !exists {
				prefix = candidate
				break
			}
		}
		ns[prefix] = name.Space
		declarations = append(declarations, declaration{prefix, name.Space})
		return prefix + ":" + name.Local, nil
	}
	qname, err := qualify(node.Name, false)
	if err != nil {
		return err
	}
	seen := map[xml.Name]bool{}
	names := make([]string, len(node.Attributes))
	for i, a := range node.Attributes {
		if seen[a.Name] {
			return fmt.Errorf("duplicate expanded inserted attribute")
		}
		seen[a.Name] = true
		if !validText(a.Value) {
			return fmt.Errorf("invalid attribute text")
		}
		names[i], err = qualify(a.Name, true)
		if err != nil {
			return err
		}
	}
	out.WriteByte('<')
	out.WriteString(qname)
	for _, decl := range declarations {
		out.WriteString(" xmlns")
		if decl.prefix != "" {
			out.WriteByte(':')
			out.WriteString(decl.prefix)
		}
		out.WriteString("=\"")
		_ = xml.EscapeText(out, []byte(decl.uri))
		out.WriteByte('"')
	}
	for i, a := range node.Attributes {
		out.WriteByte(' ')
		out.WriteString(names[i])
		out.WriteString("=\"")
		_ = xml.EscapeText(out, []byte(a.Value))
		out.WriteByte('"')
	}
	if node.Text == "" && len(node.Children) == 0 {
		out.WriteString("/>")
		return nil
	}
	out.WriteByte('>')
	_ = xml.EscapeText(out, []byte(node.Text))
	for _, child := range node.Children {
		if err := renderElement(out, child, ns, depth+1); err != nil {
			return err
		}
	}
	out.WriteString("</")
	out.WriteString(qname)
	out.WriteByte('>')
	return nil
}
