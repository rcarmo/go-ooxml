package presentation

import (
	"encoding/xml"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// EditSession provides bounded source-preserving plain text-shape corrections.
// Mutable sessions are single-owner. Layout, theme, table, group and field edits
// require later adapters and are not implied by this initial API.
type EditSession struct {
	pkg        *packaging.Preserved
	slides     map[string]bool
	generation uint64
}
type TextTarget struct {
	session          *EditSession
	generation       uint64
	doc              *losslessxml.Document
	element          losslessxml.Element
	part, hash, text string
	consumed         bool
}

func editRefusal(kind, detail string) error {
	return &packaging.Refusal{Kind: kind, Operation: "presentation_edit", Detail: detail}
}
func name(ns, local string) xml.Name { return xml.Name{Space: ns, Local: local} }
func OpenEditing(source []byte, limits packaging.Limits) (*EditSession, error) {
	p, err := packaging.OpenPreserved(source, limits)
	if err != nil {
		return nil, err
	}
	g, err := p.Graph()
	if err != nil {
		return nil, err
	}
	main := ""
	for _, e := range g.Edges {
		if e.Source == "" && e.Type == packaging.RelTypeOfficeDocument && !e.External {
			if main != "" {
				return nil, editRefusal("ambiguous_target", "multiple main parts")
			}
			main = e.ResolvedPart
		}
	}
	if main == "" {
		return nil, editRefusal("missing_target", "presentation main part missing")
	}
	types := map[string]string{}
	for _, p := range g.Parts {
		types[p.Name] = p.ContentType
	}
	if types[main] != packaging.ContentTypePresentation {
		return nil, editRefusal("unsupported_structure", "only ordinary transitional presentations supported")
	}
	b, _, err := p.Part(main)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(b)
	if err != nil {
		return nil, err
	}
	if d.Elements()[0].Name() != name(packaging.NSPresentationML, "presentation") {
		return nil, editRefusal("unsupported_structure", "presentation root expected")
	}
	rels := map[string]string{}
	for _, e := range g.Edges {
		if e.Source == main && e.Type == packaging.RelTypeSlide {
			if e.External {
				return nil, editRefusal("relationship_policy", "external slide")
			}
			rels[e.ID] = e.ResolvedPart
		}
	}
	slides := map[string]bool{}
	ids := map[string]bool{}
	for _, e := range d.Elements() {
		if e.Name() != name(packaging.NSPresentationML, "sldId") {
			continue
		}
		parent, ok := e.Parent()
		if !ok || parent.Name() != name(packaging.NSPresentationML, "sldIdLst") {
			return nil, editRefusal("unsupported_structure", "unexpected slide inventory")
		}
		rid := ""
		for _, a := range e.Attributes() {
			if a.Name == name(packaging.NSDocumentRelationships, "id") {
				rid = a.Value
			}
		}
		part, ok := rels[rid]
		if !ok || ids[rid] || slides[part] || types[part] != packaging.ContentTypeSlide {
			return nil, editRefusal("relationship_policy", "ambiguous slide identity")
		}
		ids[rid] = true
		slides[part] = true
	}
	return &EditSession{pkg: p, slides: slides}, nil
}
func (s *EditSession) FindText(part string, shapeID uint32, text string) (*TextTarget, error) {
	if !s.slides[part] {
		return nil, editRefusal("missing_target", "part is not an enrolled slide")
	}
	if text == "" {
		return nil, editRefusal("ambiguous_target", "empty text target")
	}
	b, hash, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(b)
	if err != nil {
		return nil, err
	}
	es := d.Elements()
	if es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, editRefusal("unsupported_structure", "slide root expected")
	}
	var shape losslessxml.Element
	count := 0
	for _, e := range es {
		if e.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		for _, a := range e.Attributes() {
			if a.Name == (xml.Name{Local: "id"}) {
				id, err := strconv.ParseUint(a.Value, 10, 32)
				if err != nil {
					return nil, editRefusal("unsupported_structure", "invalid shape ID")
				}
				if uint32(id) == shapeID {
					count++
					nv, ok := e.Parent()
					if !ok {
						return nil, editRefusal("unsupported_structure", "invalid shape metadata")
					}
					shape, ok = nv.Parent()
					if !ok || nv.Name() != name(packaging.NSPresentationML, "nvSpPr") || shape.Name() != name(packaging.NSPresentationML, "sp") {
						return nil, editRefusal("unsupported_structure", "target is not a plain text shape")
					}
				}
			}
		}
	}
	if count == 0 {
		return nil, editRefusal("missing_target", "shape ID absent")
	}
	if count != 1 {
		return nil, editRefusal("ambiguous_target", "duplicate shape ID")
	}
	parent, ok := shape.Parent()
	if !ok || parent.Name() != name(packaging.NSPresentationML, "spTree") {
		return nil, editRefusal("unsupported_structure", "grouped shapes require explicit group ownership")
	}
	var matches []losslessxml.Element
	for _, e := range es {
		if e.Name() != name(packaging.NSDrawingML, "t") {
			continue
		}
		got, leaf := e.Text()
		if !leaf || got != text {
			continue
		}
		for p, ok := e.Parent(); ok; p, ok = p.Parent() {
			if p == shape {
				matches = append(matches, e)
				break
			}
		}
	}
	if len(matches) == 0 {
		return nil, editRefusal("missing_target", "complete text leaf absent")
	}
	if len(matches) != 1 {
		return nil, editRefusal("ambiguous_target", "repeated exact text leaves")
	}
	return &TextTarget{session: s, generation: s.generation, doc: d, element: matches[0], part: part, hash: hash, text: text}, nil
}
func (s *EditSession) Replace(target *TextTarget, text string) error {
	if target == nil || target.session != s || target.generation != s.generation || target.consumed {
		return editRefusal("stale_target", "stale, consumed or foreign target")
	}
	_, hash, err := s.pkg.Part(target.part)
	if err != nil {
		return err
	}
	if hash != target.hash {
		return editRefusal("stale_target", "part fingerprint differs")
	}
	p := target.element
	var run, paragraph, shape losslessxml.Element
	for i, want := range []xml.Name{name(packaging.NSDrawingML, "r"), name(packaging.NSDrawingML, "p"), name(packaging.NSPresentationML, "txBody"), name(packaging.NSPresentationML, "sp")} {
		var ok bool
		p, ok = p.Parent()
		if !ok || p.Name() != want {
			return editRefusal("unsupported_structure", "text ownership is not a plain shape run")
		}
		switch i {
		case 0:
			run = p
		case 1:
			paragraph = p
		case 3:
			shape = p
		}
	}
	for _, e := range target.doc.Elements() {
		parent, ok := e.Parent()
		if !ok {
			continue
		}
		if parent == run && e != target.element && e.Name() != name(packaging.NSDrawingML, "rPr") {
			return editRefusal("unsupported_structure", "mixed run content")
		}
		if parent == paragraph && e.Name() != name(packaging.NSDrawingML, "r") && e.Name() != name(packaging.NSDrawingML, "pPr") && e.Name() != name(packaging.NSDrawingML, "endParaRPr") {
			return editRefusal("unsupported_structure", "fields/breaks/mixed paragraph")
		}
		if e.Name() == name(packaging.NSDrawingML, "spLocks") {
			for ancestor, ok := e.Parent(); ok; ancestor, ok = ancestor.Parent() {
				if ancestor == shape {
					return editRefusal("protected_operation", "locked shape requires explicit policy")
				}
			}
		}
	}
	if strings.ContainsAny(text, "\t\r\n") {
		return editRefusal("unsupported_structure", "tabs/newlines need structural authoring")
	}
	if text == target.text {
		return nil
	}
	b, err := target.doc.ReplaceText([]losslessxml.TextEdit{{Target: target.element, Text: text}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: target.part, ExpectedSHA256: target.hash, Data: b}}); err != nil {
		return err
	}
	s.generation++
	target.consumed = true
	return nil
}
func (s *EditSession) SaveAs(path string) (packaging.Receipt, error) { return s.pkg.SaveAs(path) }
