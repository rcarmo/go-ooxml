package losslessxml

import (
	"bytes"
	"encoding/xml"
	"io"
	"strings"
)

// Equivalent compares two bounded XML inputs for conservative preservation decisions.
// It does not produce canonical XML or validate an OOXML part's schema.
func Equivalent(left, right []byte) bool {
	const maxInput = 64 << 10
	if len(left) == 0 || len(right) == 0 || len(left) > maxInput || len(right) > maxInput {
		return false
	}
	// Parse validates complete roots, UTF-8, namespace bindings, duplicate attributes,
	// directives, and nesting before any equivalent result can be returned.
	if _, err := Parse(left); err != nil {
		return false
	}
	if _, err := Parse(right); err != nil {
		return false
	}
	// OPC relationship entries are identified by Id, not document order.
	// Use this narrower path only when both validated roots have the package
	// Relationships expanded name; unexpected markup fails closed.
	const relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
	root := xml.Name{Space: relationshipNamespace, Local: "Relationships"}
	ld, _ := Parse(left)
	rd, _ := Parse(right)
	if ld.Elements()[0].Name() == root || rd.Elements()[0].Name() == root {
		if ld.Elements()[0].Name() != root || rd.Elements()[0].Name() != root {
			return false
		}
		la, ok := comparisonRelationships(left, relationshipNamespace)
		if !ok {
			return false
		}
		ra, ok := comparisonRelationships(right, relationshipNamespace)
		if !ok || len(la) != len(ra) {
			return false
		}
		for id, entry := range la {
			if ra[id] != entry {
				return false
			}
		}
		return true
	}
	a, b := xml.NewDecoder(bytes.NewReader(left)), xml.NewDecoder(bytes.NewReader(right))
	as, bs := []map[string]string{{"xml": xmlNS}}, []map[string]string{{"xml": xmlNS}}
	for {
		astart, bstart := a.InputOffset(), b.InputOffset()
		at, ae := a.Token()
		bt, be := b.Token()
		if ae == io.EOF || be == io.EOF {
			return ae == io.EOF && be == io.EOF
		}
		if ae != nil || be != nil {
			return false
		}
		switch x := at.(type) {
		case xml.StartElement:
			y, ok := bt.(xml.StartElement)
			if !ok || x.Name != y.Name {
				return false
			}
			an := comparisonScope(as[len(as)-1], x.Attr)
			bn := comparisonScope(bs[len(bs)-1], y.Attr)
			if !comparisonAttrs(x.Attr, y.Attr, an, bn) {
				return false
			}
			as, bs = append(as, an), append(bs, bn)
		case xml.EndElement:
			y, ok := bt.(xml.EndElement)
			if !ok || x.Name != y.Name || len(as) <= 1 || len(bs) <= 1 {
				return false
			}
			as, bs = as[:len(as)-1], bs[:len(bs)-1]
		case xml.CharData:
			y, ok := bt.(xml.CharData)
			if !ok || !bytes.Equal(x, y) || !comparisonBoundReferences(string(x), as[len(as)-1], bs[len(bs)-1]) {
				return false
			}
		case xml.Comment:
			y, ok := bt.(xml.Comment)
			if !ok || !bytes.Equal(x, y) {
				return false
			}
		case xml.ProcInst:
			y, ok := bt.(xml.ProcInst)
			// Token() normalises the whitespace after a PI target. Keep its
			// source spelling significant rather than erasing that difference.
			if !ok || x.Target != y.Target || !bytes.Equal(x.Inst, y.Inst) || !bytes.Equal(left[astart:a.InputOffset()], right[bstart:b.InputOffset()]) {
				return false
			}
		default:
			return false // directives and unsupported token types fail closed
		}
	}
}

// Only ordinary OPC relationship entries are reorderable. Require all three
// attributes and unique IDs; comments, processing instructions, mixed content,
// extensions and duplicate/unknown attributes cannot be erased by sorting.
func comparisonRelationships(source []byte, namespace string) (map[string][3]string, bool) {
	d := xml.NewDecoder(bytes.NewReader(source))
	entries := map[string][3]string{}
	depth := 0
	seenRoot := false
	for {
		token, err := d.Token()
		if err == io.EOF {
			return entries, seenRoot && depth == 0
		}
		if err != nil {
			return nil, false
		}
		switch x := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if seenRoot || x.Name != (xml.Name{Space: namespace, Local: "Relationships"}) {
					return nil, false
				}
				seenRoot = true
				for _, a := range x.Attr {
					if !comparisonNamespaceDeclaration(a) {
						return nil, false
					}
				}
				continue
			}
			if depth != 2 || x.Name != (xml.Name{Space: namespace, Local: "Relationship"}) {
				return nil, false
			}
			values := map[string]string{}
			for _, a := range x.Attr {
				if comparisonNamespaceDeclaration(a) {
					continue
				}
				if a.Name.Space != "" || (a.Name.Local != "Id" && a.Name.Local != "Type" && a.Name.Local != "Target" && a.Name.Local != "TargetMode") {
					return nil, false
				}
				if _, exists := values[a.Name.Local]; exists {
					return nil, false
				}
				values[a.Name.Local] = a.Value
			}
			id, hasID := values["Id"]
			_, hasType := values["Type"]
			_, hasTarget := values["Target"]
			if !hasID || id == "" || !hasType || values["Type"] == "" || !hasTarget || values["Target"] == "" {
				return nil, false
			}
			if _, exists := entries[id]; exists {
				return nil, false
			}
			entries[id] = [3]string{values["Type"], values["Target"], values["TargetMode"]}
		case xml.EndElement:
			depth--
			if depth < 0 {
				return nil, false
			}
		case xml.CharData:
			if len(x) != 0 {
				return nil, false
			}
		default:
			return nil, false
		}
	}
}

func comparisonNamespaceDeclaration(a xml.Attr) bool {
	return a.Name.Space == "xmlns" || a.Name.Space == "" && a.Name.Local == "xmlns"
}

func comparisonScope(parent map[string]string, attrs []xml.Attr) map[string]string {
	scope := make(map[string]string, len(parent)+len(attrs))
	for k, v := range parent {
		scope[k] = v
	}
	for _, a := range attrs {
		if a.Name.Space == "xmlns" {
			scope[a.Name.Local] = a.Value
		} else if a.Name.Space == "" && a.Name.Local == "xmlns" {
			scope[""] = a.Value
		}
	}
	return scope
}

func comparisonAttrs(left, right []xml.Attr, ln, rn map[string]string) bool {
	values := make(map[xml.Name]string, len(left))
	for _, a := range left {
		if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			continue
		}
		values[a.Name] = a.Value
	}
	count := 0
	for _, a := range right {
		if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			continue
		}
		v, ok := values[a.Name]
		if !ok || !comparisonValue(v, a.Value, ln, rn) {
			return false
		}
		if a.Name.Space == "http://schemas.openxmlformats.org/markup-compatibility/2006" && a.Name.Local == "Ignorable" && !comparisonPrefixList(a.Value, ln, rn) {
			return false
		}
		count++
	}
	return count == len(values)
}

// Attribute values retain lexical spelling. An in-scope prefix reference can
// occur in text, a QName-like value or a whitespace-separated prefix list.
// Without schema types, this may reject ordinary colon-bearing values.
func comparisonValue(left, right string, ln, rn map[string]string) bool {
	return left == right && comparisonBoundReferences(left, ln, rn)
}

func comparisonBoundReferences(value string, ln, rn map[string]string) bool {
	for _, word := range strings.Fields(value) {
		if comparisonQNameLike(word) {
			prefix, _, _ := strings.Cut(word, ":")
			l, lok := ln[prefix]
			r, rok := rn[prefix]
			if !lok || !rok || l == "" || l != r {
				return false
			}
			continue
		}
		if localName(word) {
			l, lok := ln[word]
			r, rok := rn[word]
			if lok || rok {
				if !lok || !rok || l == "" || l != r {
					return false
				}
			}
		} else if strings.Contains(word, ":") && !comparisonSameScope(ln, rn) {
			// Unknown colon-bearing token spelling under different bindings.
			return false
		}
	}
	return true
}

// Ignorable carries space-separated prefixes; unlike an ordinary simple
// attribute value, every listed prefix must resolve in both inputs.
func comparisonPrefixList(value string, left, right map[string]string) bool {
	for _, prefix := range strings.Fields(value) {
		if !localName(prefix) || left[prefix] == "" || left[prefix] != right[prefix] {
			return false
		}
	}
	return true
}

func comparisonSameScope(left, right map[string]string) bool {
	if len(left) != len(right) {
		return false
	}
	for k, v := range left {
		if right[k] != v {
			return false
		}
	}
	return true
}

func comparisonQNameLike(s string) bool {
	if strings.Count(s, ":") != 1 {
		return false
	}
	prefix, local, _ := strings.Cut(s, ":")
	return localName(prefix) && localName(local)
}
