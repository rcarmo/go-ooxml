package presentation

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ContractSlideNotes is a read-only observation of one relationship-ordered slide.
// Visible field text is returned as stored; fields are never evaluated.
type ContractSlideNotes struct {
	Part     string
	Title    string
	HasNotes bool
	Text     string
}

// ReadContractSlideNotes inspects the enrolled slide order and related existing
// notes without authoring a part. Unknown/ambiguous ownership refuses.
func ReadContractSlideNotes(source []byte) ([]ContractSlideNotes, error) {
	s, err := OpenEditing(source, packaging.Limits{})
	if err != nil {
		return nil, err
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return nil, err
	}
	main := ""
	for _, e := range graph.Edges {
		if e.Source == "" && e.Type == packaging.RelTypeOfficeDocument && !e.External {
			if main != "" {
				return nil, editRefusal("PPTX_PRESENTATION_INVALID", "ambiguous presentation")
			}
			main = e.ResolvedPart
		}
	}
	if main == "" {
		return nil, editRefusal("PPTX_PRESENTATION_INVALID", "presentation absent")
	}
	pres, _, err := s.pkg.Part(main)
	if err != nil {
		return nil, err
	}
	doc, err := losslessxml.Parse(pres)
	if err != nil {
		return nil, err
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	rName := xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}
	if len(doc.Elements()) == 0 || doc.Elements()[0].Name() != p("presentation") {
		return nil, editRefusal("PPTX_PRESENTATION_INVALID", "presentation root")
	}
	rels := map[string]string{}
	for _, e := range graph.Edges {
		if e.Source == main && e.Type == packaging.RelTypeSlide {
			if e.External || rels[e.ID] != "" {
				return nil, editRefusal("PPTX_PRESENTATION_INVALID", "ambiguous slide edge")
			}
			rels[e.ID] = e.ResolvedPart
		}
	}
	result := []ContractSlideNotes{}
	seen := map[string]bool{}
	ids := map[string]bool{}
	for _, n := range doc.Elements() {
		if n.Name() != p("sldId") {
			continue
		}
		owner, ok := n.Parent()
		if !ok || owner.Name() != p("sldIdLst") {
			return nil, editRefusal("PPTX_PRESENTATION_INVALID", "unexpected slide list")
		}
		rid, id := "", ""
		for _, attr := range n.Attributes() {
			if attr.Name == rName {
				rid = attr.Value
			}
			if attr.Name.Space == "" && attr.Name.Local == "id" {
				id = attr.Value
			}
		}
		value, e := strconv.ParseUint(id, 10, 32)
		part := rels[rid]
		if e != nil || value == 0 || part == "" || seen[part] || ids[id] || !s.slides[part] {
			return nil, editRefusal("PPTX_PRESENTATION_INVALID", "invalid slide identity")
		}
		seen[part] = true
		ids[id] = true
		slide, _, e := s.pkg.Part(part)
		if e != nil {
			return nil, e
		}
		title, e := contractVisiblePlaceholder(slide, "title", "ctrTitle")
		if e != nil {
			return nil, e
		}
		record := ContractSlideNotes{Part: part, Title: title}
		notesPart := ""
		for _, edge := range graph.Edges {
			if edge.Source == part && edge.Type == packaging.RelTypeNotesSlide {
				if edge.External || notesPart != "" {
					return nil, editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "ambiguous notes ownership")
				}
				notesPart = edge.ResolvedPart
			}
		}
		if notesPart != "" {
			// The relationship alone does not make an arbitrary XML member a
			// notes slide. Resolve its declared OPC role before reading text.
			role, inbound := "", 0
			for _, entry := range graph.Parts {
				if entry.Name == notesPart {
					role, inbound = entry.ContentType, entry.Inbound
					break
				}
			}
			if role != packaging.ContentTypeNotesSlide || inbound != 1 {
				return nil, editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "notes part ownership or content type")
			}
			record.HasNotes = true
			notes, _, e := s.pkg.Part(notesPart)
			if e != nil {
				return nil, e
			}
			record.Text, e = contractVisibleNotes(notes)
			if e != nil {
				return nil, e
			}
		}
		result = append(result, record)
	}
	if len(result) == 0 || len(result) != len(rels) {
		return nil, editRefusal("PPTX_PRESENTATION_INVALID", "slide inventory differs from relationships")
	}
	return result, nil
}

func contractVisiblePlaceholder(data []byte, kinds ...string) (string, error) {
	d, err := losslessxml.Parse(data)
	if err != nil {
		return "", err
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	if len(d.Elements()) == 0 || d.Elements()[0].Name() != p("sld") {
		return "", editRefusal("PPTX_PRESENTATION_INVALID", "slide root")
	}
	var selected losslessxml.Element
	count := 0
	for _, n := range d.Elements() {
		if n.Name() != p("ph") {
			continue
		}
		for _, attr := range n.Attributes() {
			if attr.Name.Space != "" || attr.Name.Local != "type" {
				continue
			}
			for _, kind := range kinds {
				if attr.Value == kind {
					for parent, ok := n.Parent(); ok; parent, ok = parent.Parent() {
						if parent.Name() == p("sp") {
							selected = parent
							count++
							break
						}
					}
				}
			}
		}
	}
	if count == 0 {
		// An ordinary blank slide with no shape or visible text has no title.
		// Other content without a title requires a separate interpretation.
		var ordinary bool
		for _, n := range d.Elements() {
			if n.Name() == p("sp") || n.Name() == p("graphicFrame") || n.Name() == p("pic") || n.Name() == p("grpSp") || n.Name() == p("ph") || n.Name() == (xml.Name{Space: packaging.NSDrawingML, Local: "t"}) {
				ordinary = true
			}
		}
		if !ordinary {
			return "", nil
		}
	}
	if count != 1 {
		return "", editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", fmt.Sprintf("title owner count %d", count))
	}
	var text strings.Builder
	for _, n := range d.Elements() {
		if n.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: "t"}) || !contractDescendant(n, selected) {
			continue
		}
		v, leaf := n.Text()
		if !leaf {
			return "", editRefusal("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "non-leaf title")
		}
		text.WriteString(v)
	}
	return text.String(), nil
}

func contractDescendant(child, owner losslessxml.Element) bool {
	for n, ok := child.Parent(); ok; n, ok = n.Parent() {
		if n == owner {
			return true
		}
	}
	return false
}

func contractVisibleNotes(data []byte) (string, error) {
	d, err := losslessxml.Parse(data)
	if err != nil {
		return "", err
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	if len(d.Elements()) == 0 || d.Elements()[0].Name() != p("notes") {
		return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "notes root")
	}
	var shape losslessxml.Element
	count := 0
	for _, n := range d.Elements() {
		if n.Name() != p("ph") {
			continue
		}
		for _, attr := range n.Attributes() {
			if attr.Name.Space == "" && attr.Name.Local == "type" && attr.Value == "body" {
				for ancestor, ok := n.Parent(); ok; ancestor, ok = ancestor.Parent() {
					if ancestor.Name() == p("sp") {
						shape = ancestor
						count++
						break
					}
				}
			}
		}
	}
	if count != 1 {
		return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "notes body ambiguous")
	}
	var body losslessxml.Element
	bodyCount := 0
	for _, n := range d.Elements() {
		if n.Name() == p("txBody") {
			owner, ok := n.Parent()
			if ok && owner == shape {
				body = n
				bodyCount++
			}
		}
	}
	if bodyCount != 1 {
		return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "notes text body ambiguous")
	}
	// Validate the body subtree before projecting it. A foreign child of a
	// run/field or of the body must not disappear from the visible result.
	for _, node := range d.Elements() {
		if !contractDescendant(node, body) {
			continue
		}
		parent, ok := node.Parent()
		if !ok {
			return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unowned notes child")
		}
		switch parent.Name() {
		case p("txBody"):
			if node.Name() != a("bodyPr") && node.Name() != a("lstStyle") && node.Name() != a("p") {
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes body child")
			}
		case a("p"):
			if node.Name() != a("pPr") && node.Name() != a("endParaRPr") && node.Name() != a("br") && node.Name() != a("r") && node.Name() != a("fld") {
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes paragraph child")
			}
		case a("r"):
			if node.Name() != a("rPr") && node.Name() != a("t") {
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes run child")
			}
		case a("fld"):
			if node.Name() != a("rPr") && node.Name() != a("pPr") && node.Name() != a("t") {
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes field child")
			}
		case a("rPr"):
			// Only this bounded formatting structure is understood. An
			// arbitrary descendant could hide visible content from projection.
			if node.Name() != a("solidFill") && node.Name() != a("latin") {
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes run property")
			}
		case a("solidFill"):
			if node.Name() != a("srgbClr") {
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes fill property")
			}
		default:
			// All other formatting elements are opaque leaves, not a way to
			// smuggle text, breaks, fields, or unknown namespace content.
			return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes property descendant")
		}
	}
	paragraphs := []string{}
	for _, paragraph := range d.Elements() {
		if paragraph.Name() != a("p") {
			continue
		}
		owner, ok := paragraph.Parent()
		if !ok || owner != body {
			continue
		}
		var line strings.Builder
		for _, node := range d.Elements() {
			parent, ok := node.Parent()
			if !ok || parent != paragraph {
				continue
			}
			switch node.Name() {
			case a("pPr"), a("endParaRPr"):
				continue
			case a("br"):
				line.WriteByte('\n')
			case a("r"), a("fld"):
				textCount := 0
				for _, child := range d.Elements() {
					direct, ok := child.Parent()
					if !ok || direct != node || child.Name() != a("t") {
						continue
					}
					v, leaf := child.Text()
					if !leaf {
						return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "non-leaf visible text")
					}
					textCount++
					line.WriteString(contractNormaliseVisibleNotesBreaks(v))
				}
				if textCount != 1 {
					return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "visible notes fragment missing")
				}
			default:
				return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "unsupported notes paragraph child")
			}
		}
		paragraphs = append(paragraphs, line.String())
	}
	if len(paragraphs) == 0 {
		return "", editRefusal("PPTX_NOTES_STRUCTURE_UNSUPPORTED", "notes paragraphs absent")
	}
	return strings.Join(paragraphs, "\n"), nil
}

func contractNormaliseVisibleNotesBreaks(text string) string {
	return strings.ReplaceAll(strings.ReplaceAll(text, "\r\n", "\n"), "\r", "\n")
}
