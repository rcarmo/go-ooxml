package presentation

import (
	"bytes"
	"fmt"
	"regexp"
	"sort"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Independent lexical masks restrict each style adapter to its declared spans.
func graphicsMaskStyle(t *testing.T, source []byte, id uint32, profile string) []byte {
	t.Helper()
	doc, e := losslessxml.Parse(source)
	if e != nil {
		t.Fatal(e)
	}
	var shape losslessxml.Element
	for _, n := range doc.Elements() {
		nv := ""
		if n.Name() == name(packaging.NSPresentationML, "sp") {
			nv = "nvSpPr"
		}
		if n.Name() == name(packaging.NSPresentationML, "pic") {
			nv = "nvPicPr"
		}
		if n.Name() == name(packaging.NSPresentationML, "cxnSp") {
			nv = "nvCxnSpPr"
		}
		if nv == "" {
			continue
		}
		got, _, _, e := graphicsIdentity(doc, n, nv)
		if e != nil {
			t.Fatal(e)
		}
		if got == id {
			shape = n
			break
		}
	}
	if shape.Name().Local == "" {
		t.Fatal("custody target absent")
	}
	type edit struct {
		a, z  int
		value []byte
	}
	edits := []edit{}
	switch profile {
	case "gradient":
		pr, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
		if e != nil {
			t.Fatal(e)
		}
		for _, n := range graphicsChildren(doc, pr) {
			if n.Name().Space == packaging.NSDrawingML && graphicsChoice(n.Name().Local, []string{"noFill", "solidFill", "gradFill"}) {
				a, z := n.SourceRange()
				edits = append(edits, edit{a, z, nil})
			}
		}
	case "outline":
		pr, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
		if e != nil {
			t.Fatal(e)
		}
		line, e := graphicsOne(doc, pr, packaging.NSDrawingML, "ln", false)
		if e != nil {
			t.Fatal(e)
		}
		a, _ := line.SourceRange()
		z, _ := line.ContentRange()
		// Strip only writable scalar attributes; width, alignment, bindings and
		// other opening-tag bytes must remain literal too.
		opening := bytes.Clone(source[a:z])
		opening = regexp.MustCompile(`\s+(?:cap|cmpd)\s*=\s*(?:"[^"]*"|'[^']*')`).ReplaceAll(opening, nil)
		edits = append(edits, edit{a, z, opening})
		for _, n := range graphicsChildren(doc, line) {
			if graphicsChoice(n.Name().Local, []string{"round", "bevel", "miter", "headEnd", "tailEnd"}) {
				a, z := n.SourceRange()
				edits = append(edits, edit{a, z, nil})
			}
		}
	case "shape-alpha", "picture-alpha":
		var owner losslessxml.Element
		for _, n := range doc.Elements() {
			if !manipulationWithin(n, shape) {
				continue
			}
			if profile == "picture-alpha" && n.Name() == name(packaging.NSDrawingML, "blip") {
				owner = n
				break
			}
			if profile == "shape-alpha" && (n.Name() == name(packaging.NSDrawingML, "srgbClr") || n.Name() == name(packaging.NSDrawingML, "schemeClr")) {
				parent, ok := n.Parent()
				if ok && parent.Name() == name(packaging.NSDrawingML, "solidFill") {
					owner = n
					break
				}
			}
		}
		if owner.Name().Local == "" {
			t.Fatal("alpha owner absent")
		}
		tag := "alpha"
		if profile == "picture-alpha" {
			tag = "alphaModFix"
		}
		var removed []losslessxml.Element
		for _, n := range graphicsChildren(doc, owner) {
			if n.Name() == name(packaging.NSDrawingML, tag) {
				removed = append(removed, n)
			}
		}
		raw := owner.Raw()
		start, _ := owner.SourceRange()
		for i := len(removed) - 1; i >= 0; i-- {
			a, z := removed[i].SourceRange()
			raw = append(append([]byte{}, raw[:a-start]...), raw[z-start:]...)
		}
		// Only an empty owner's self-closing/expanded form may differ on insertion.
		if len(graphicsChildren(doc, owner)) == len(removed) {
			opening := fmtOpeningEnd(raw)
			if opening < 1 {
				t.Fatal("owner lexical tag")
			}
			raw = append([]byte{}, raw[:opening]...)
			if bytes.HasSuffix(raw, []byte("/>")) {
				raw = append(raw[:len(raw)-2], '>')
			}
			raw = append(raw, []byte("</"+owner.QualifiedName()+">")...)
		}
		a, z := owner.SourceRange()
		edits = append(edits, edit{a, z, raw})
	default:
		t.Fatal("unknown custody profile", profile)
	}
	sort.Slice(edits, func(i, j int) bool { return edits[i].a > edits[j].a })
	result := bytes.Clone(source)
	for _, edit := range edits {
		result = append(append(append([]byte{}, result[:edit.a]...), edit.value...), result[edit.z:]...)
	}
	return result
}
func graphicsAssertStyleCustody(t *testing.T, before, after []byte, id uint32, profile string) {
	t.Helper()
	a, b := graphicsMaskStyle(t, before, id, profile), graphicsMaskStyle(t, after, id, profile)
	if !bytes.Equal(a, b) {
		t.Fatal(fmt.Sprintf("XML changed outside %s spans", profile))
	}
}
