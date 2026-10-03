package presentation

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Admit only directly owned, ordinary title-layout placeholder shapes. A
// descendant p:ph in extLst, group, alternate content or foreign nv metadata
// is not an authored placeholder; it must refuse rather than count as one.
func contractAppendLayoutPlaceholders(doc *losslessxml.Document) (titleKind, subtitleIndex string, err error) {
	unsafe := func(why string) (string, string, error) { return "", "", fmt.Errorf("unsafe title layout: %s", why) }
	p := func(local string) xml.Name { return name(packaging.NSPresentationML, local) }
	a := func(local string) xml.Name { return name(packaging.NSDrawingML, local) }
	elements := doc.Elements()
	if len(elements) == 0 || elements[0].Name() != p("sldLayout") {
		return unsafe("layout root")
	}
	children := func(parent losslessxml.Element, kind xml.Name) []losslessxml.Element {
		return contractTitleChildren(doc, parent, kind)
	}
	only := func(parent losslessxml.Element, kind xml.Name) (losslessxml.Element, bool) {
		items := children(parent, kind)
		if len(items) != 1 {
			return losslessxml.Element{}, false
		}
		return items[0], true
	}
	common, ok := only(elements[0], p("cSld"))
	if !ok {
		return unsafe("unique common slide")
	}
	tree, ok := only(common, p("spTree"))
	if !ok {
		return unsafe("unique shape tree")
	}
	// Duplicate object IDs anywhere in the layout can make placeholder
	// inheritance ambiguous even when the selected shape itself is unique.
	identityCounts := map[uint64]int{}
	for _, node := range elements {
		if node.Name() != p("cNvPr") {
			continue
		}
		idCount := 0
		var id uint64
		for _, attr := range node.Attributes() {
			if attr.Name == (xml.Name{Local: "id"}) {
				n, e := strconv.ParseUint(attr.Value, 10, 32)
				if e != nil || n == 0 {
					return unsafe("invalid layout shape identity")
				}
				id = n
				idCount++
			}
		}
		if idCount != 1 {
			return unsafe("missing/duplicate layout shape ID")
		}
		identityCounts[id]++
		if identityCounts[id] != 1 {
			return unsafe("duplicate layout shape ID")
		}
	}
	seenTitle, seenSubtitle := 0, 0
	for _, node := range elements {
		if node.Name() != p("ph") {
			continue
		}
		attrs := node.Attributes()
		kind, idx := "", ""
		hasKind, hasIdx := false, false
		for _, attr := range attrs {
			switch attr.Name {
			case xml.Name{Local: "type"}:
				if hasKind {
					return unsafe("duplicate placeholder type")
				}
				kind, hasKind = attr.Value, true
			case xml.Name{Local: "idx"}:
				if hasIdx {
					return unsafe("duplicate placeholder index")
				}
				idx, hasIdx = attr.Value, true
			}
		}
		if kind != "ctrTitle" && kind != "title" && kind != "subTitle" {
			continue
		}
		if !hasKind || len(attrs) != 1+boolToInt(hasIdx) {
			return unsafe("placeholder attributes")
		}
		nv, ok := node.Parent()
		if !ok || nv.Name() != p("nvPr") {
			return unsafe("placeholder owner")
		}
		metadata, ok := nv.Parent()
		if !ok || metadata.Name() != p("nvSpPr") {
			return unsafe("placeholder metadata owner")
		}
		shape, ok := metadata.Parent()
		if !ok || shape.Name() != p("sp") {
			return unsafe("placeholder shape owner")
		}
		owner, ok := shape.Parent()
		if !ok || owner != tree {
			return unsafe("placeholder outside direct shape tree")
		}
		if direct, ok := only(metadata, p("nvPr")); !ok || direct != nv || len(children(nv, p("ph"))) != 1 {
			return unsafe("ambiguous placeholder metadata")
		}
		identity, ok := only(metadata, p("cNvPr"))
		if !ok {
			return unsafe("missing shape identity")
		}
		if _, ok = only(metadata, p("cNvSpPr")); !ok {
			return unsafe("missing shape properties")
		}
		if _, ok = only(shape, p("spPr")); !ok {
			return unsafe("missing shape geometry")
		}
		body, ok := only(shape, p("txBody"))
		if !ok {
			return unsafe("missing shape text body")
		}
		if _, ok = only(body, a("bodyPr")); !ok {
			return unsafe("missing text body properties")
		}
		if _, ok = only(body, a("lstStyle")); !ok {
			return unsafe("missing text list style")
		}
		if len(children(body, a("p"))) == 0 {
			return unsafe("missing text paragraph")
		}
		id := uint64(0)
		idCount := 0
		for _, attr := range identity.Attributes() {
			if attr.Name == (xml.Name{Local: "id"}) {
				n, e := strconv.ParseUint(attr.Value, 10, 32)
				if e != nil || n == 0 {
					return unsafe("invalid shape identity")
				}
				id = n
				idCount++
			}
		}
		if idCount != 1 || identityCounts[id] != 1 {
			return unsafe("ambiguous shape identity")
		}
		switch kind {
		case "ctrTitle", "title":
			if hasIdx {
				return unsafe("indexed title placeholder")
			}
			seenTitle++
			titleKind = kind
		case "subTitle":
			if !hasIdx {
				return unsafe("subtitle index required")
			}
			n, e := strconv.ParseUint(idx, 10, 32)
			if e != nil || n == 0 || n > 9999 || strconv.FormatUint(n, 10) != idx {
				return unsafe("bounded canonical subtitle index required")
			}
			seenSubtitle++
			subtitleIndex = idx
		}
	}
	if seenTitle != 1 || seenSubtitle != 1 {
		return unsafe("unique title and subtitle shapes required")
	}
	// Every selected placeholder was observed through the ownership chain above;
	// the source layout stays byte-for-byte untouched by append.
	return titleKind, subtitleIndex, nil
}
func boolToInt(v bool) int {
	if v {
		return 1
	}
	return 0
}
