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
			if !ok || !bytes.Equal(x, y) {
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
		count++
	}
	return count == len(values)
}

// Attribute values retain lexical spelling. Identical QName-like values with
// different or unbound prefix scopes are conservatively non-equivalent.
// Attribute schema typing is unknown, so aliases in values cannot be normalised.
func comparisonValue(left, right string, ln, rn map[string]string) bool {
	if left != right {
		return false
	}
	if !comparisonQNameLike(left) {
		return true
	}
	lp, ll, lok := comparisonQName(left, ln)
	rp, rl, rok := comparisonQName(right, rn)
	return lok && rok && lp == rp && ll == rl
}

func comparisonQName(s string, ns map[string]string) (uri, local string, ok bool) {
	if !comparisonQNameLike(s) {
		return "", "", false
	}
	prefix, local, _ := strings.Cut(s, ":")
	uri, ok = ns[prefix]
	return uri, local, ok && uri != ""
}

func comparisonQNameLike(s string) bool {
	if strings.Count(s, ":") != 1 {
		return false
	}
	prefix, local, _ := strings.Cut(s, ":")
	return localName(prefix) && localName(local)
}
