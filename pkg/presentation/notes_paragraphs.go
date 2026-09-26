package presentation

import (
	"bytes"
	"encoding/xml"
	"io"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Explicitly supported formatting templates are copied by expanded name. QName-
// valued extensions and relationship-bearing properties never reach authoring.
var notesPropertyAttrs = map[string]string{
	"pPr":        "marL marR lvl indent algn defTabSz rtl eaLnBrk fontAlgn latinLnBrk hangingPunct",
	"rPr":        "lang altLang sz b i u strike kern cap spc normalizeH baseline noProof dirty err smtClean",
	"defRPr":     "lang altLang sz b i u strike kern cap spc normalizeH baseline noProof dirty err smtClean",
	"endParaRPr": "lang altLang sz b i u strike kern cap spc normalizeH baseline noProof dirty err smtClean",
	"lnSpc":      "", "spcBef": "", "spcAft": "", "spcPct": "val", "spcPts": "val", "solidFill": "", "noFill": "",
	"srgbClr": "val", "schemeClr": "val", "latin": "typeface pitchFamily charset", "ea": "typeface pitchFamily charset", "cs": "typeface pitchFamily charset",
	"buNone": "", "buChar": "char", "buAutoNum": "type startAt", "buFont": "typeface pitchFamily charset", "buSzPct": "val", "buSzPts": "val", "buClr": "", "buClrTx": "", "buSzTx": "", "buFontTx": "",
}
var notesPropertyChildren = map[string]string{
	"pPr": "lnSpc spcBef spcAft buClrTx buClr buSzTx buSzPct buSzPts buFontTx buFont buNone buAutoNum buChar defRPr",
	"rPr": "noFill solidFill latin ea cs", "defRPr": "noFill solidFill latin ea cs", "endParaRPr": "noFill solidFill latin ea cs",
	"lnSpc": "spcPct spcPts", "spcBef": "spcPct spcPts", "spcAft": "spcPct spcPts", "solidFill": "srgbClr schemeClr", "buClr": "srgbClr schemeClr",
}

func notesTokenIn(list, value string) bool {
	for _, v := range strings.Fields(list) {
		if v == value {
			return true
		}
	}
	return false
}
func notesProperty(doc *losslessxml.Document, e losslessxml.Element) (losslessxml.NewElement, error) {
	n := e.Name()
	allowed, ok := notesPropertyAttrs[n.Local]
	if !ok || n.Space != packaging.NSDrawingML {
		return losslessxml.NewElement{}, editRefusal("unsupported_structure", "unproved notes formatting property")
	}
	node := losslessxml.NewElement{Name: n, Attributes: e.Attributes()}
	for _, a := range node.Attributes {
		if a.Name.Space != "" || !notesTokenIn(allowed, a.Name.Local) {
			return node, editRefusal("unsupported_structure", "unproved notes formatting attribute")
		}
	}
	text, _ := e.Text()
	if strings.TrimSpace(text) != "" {
		return node, editRefusal("unsupported_structure", "mixed notes formatting text")
	}
	seen := map[string]bool{}
	for _, child := range doc.Elements() {
		if p, ok := child.Parent(); !ok || p != e {
			continue
		}
		if !notesTokenIn(notesPropertyChildren[n.Local], child.Name().Local) || seen[child.Name().Local] {
			return node, editRefusal("unsupported_structure", "unproved/duplicate notes formatting child")
		}
		seen[child.Name().Local] = true
		c, err := notesProperty(doc, child)
		if err != nil {
			return node, err
		}
		node.Children = append(node.Children, c)
	}
	return node, nil
}
func notesPlainMarkup(e losslessxml.Element) error {
	d := xml.NewDecoder(bytes.NewReader(e.Raw()))
	for {
		token, err := d.RawToken()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return editRefusal("unsupported_structure", err.Error())
		}
		switch token.(type) {
		case xml.Comment, xml.ProcInst, xml.Directive:
			return editRefusal("unsupported_structure", "opaque markup inside notes paragraph")
		}
	}
}
func readNotesParagraphs(doc *losslessxml.Document, body losslessxml.Element, target *NotesTarget) error {
	target.paragraphs = notesChildren(doc, body, name(packaging.NSDrawingML, "p"))
	if len(target.paragraphs) == 0 {
		return editRefusal("unsupported_structure", "notes body requires an existing paragraph")
	}
	for _, e := range doc.Elements() {
		if p, ok := e.Parent(); ok && p == body && e.Name() != name(packaging.NSDrawingML, "p") && e.Name() != name(packaging.NSDrawingML, "bodyPr") && e.Name() != name(packaging.NSDrawingML, "lstStyle") {
			return editRefusal("unsupported_structure", "unknown notes body child")
		}
	}
	lines := []string{}
	leafCount := 0
	for pi, paragraph := range target.paragraphs {
		if len(paragraph.Attributes()) != 0 {
			return editRefusal("unsupported_structure", "unknown notes paragraph attributes")
		}
		if err := notesPlainMarkup(paragraph); err != nil {
			return err
		}
		if text, _ := paragraph.Text(); strings.TrimSpace(text) != "" {
			return editRefusal("unsupported_structure", "mixed notes paragraph text")
		}
		var line strings.Builder
		seenPPr, seenEnd, seenRun := false, false, false
		for _, e := range doc.Elements() {
			parent, ok := e.Parent()
			if !ok || parent != paragraph {
				continue
			}
			n := e.Name()
			if n.Space != packaging.NSDrawingML {
				return editRefusal("unsupported_structure", "unknown notes paragraph namespace")
			}
			switch n.Local {
			case "pPr", "endParaRPr":
				if n.Local == "pPr" {
					if seenPPr || seenRun || seenEnd {
						return editRefusal("unsupported_structure", "duplicate/misplaced notes paragraph properties")
					}
					seenPPr = true
				} else {
					if seenEnd {
						return editRefusal("unsupported_structure", "duplicate notes end properties")
					}
					seenEnd = true
				}
				property, err := notesProperty(doc, e)
				if err != nil {
					return err
				}
				if pi == 0 {
					if n.Local == "pPr" {
						target.pPr = &property
					} else {
						target.endPr = &property
					}
				}
			case "r":
				if seenEnd || len(e.Attributes()) != 0 {
					return editRefusal("unsupported_structure", "misplaced or extended notes run")
				}
				if text, _ := e.Text(); strings.TrimSpace(text) != "" {
					return editRefusal("unsupported_structure", "mixed notes run text")
				}
				firstRun := !seenRun
				seenRun = true
				hasText, hasPr := false, false
				for _, child := range doc.Elements() {
					if p, ok := child.Parent(); !ok || p != e {
						continue
					}
					switch child.Name() {
					case name(packaging.NSDrawingML, "rPr"):
						if hasPr || hasText {
							return editRefusal("unsupported_structure", "duplicate/misplaced run properties")
						}
						hasPr = true
						property, err := notesProperty(doc, child)
						if err != nil {
							return err
						}
						if pi == 0 && firstRun {
							target.rPr = &property
						}
					case name(packaging.NSDrawingML, "t"):
						if hasText {
							return editRefusal("unsupported_structure", "duplicate notes text leaf")
						}
						hasText = true
						text, leaf := child.Text()
						if !leaf || strings.ContainsAny(text, "\r\n\t") {
							return editRefusal("unsupported_structure", "nonordinary notes text leaf")
						}
						for _, a := range child.Attributes() {
							if a.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || (a.Value != "default" && a.Value != "preserve") {
								return editRefusal("unsupported_structure", "unknown notes text attribute")
							}
						}
						line.WriteString(text)
						leafCount++
						target.leaf = child
					default:
						return editRefusal("unsupported_structure", "field/break/unknown notes run content")
					}
				}
				if !hasText {
					return editRefusal("unsupported_structure", "notes run lacks text")
				}
			default:
				return editRefusal("unsupported_structure", "field/break/unknown notes paragraph child")
			}
		}
		lines = append(lines, line.String())
	}
	target.text = strings.Join(lines, "\n")
	target.singleLeaf = len(target.paragraphs) == 1 && leafCount == 1
	return nil
}
func notesNewParagraph(target *NotesTarget, line string) losslessxml.NewElement {
	p := losslessxml.NewElement{Name: name(packaging.NSDrawingML, "p")}
	if target.pPr != nil {
		p.Children = append(p.Children, *target.pPr)
	}
	if line != "" {
		r := losslessxml.NewElement{Name: name(packaging.NSDrawingML, "r")}
		if target.rPr != nil {
			r.Children = append(r.Children, *target.rPr)
		}
		text := losslessxml.NewElement{Name: name(packaging.NSDrawingML, "t"), Text: line}
		if strings.Trim(line, " ") != line {
			text.Attributes = []xml.Attr{{Name: xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}, Value: "preserve"}}
		}
		r.Children = append(r.Children, text)
		p.Children = append(p.Children, r)
	}
	if target.endPr != nil {
		p.Children = append(p.Children, *target.endPr)
	}
	return p
}
