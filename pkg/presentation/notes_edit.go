package presentation

import (
	"encoding/xml"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// NotesTarget identifies an existing plain notes body in one immutable snapshot.
// Text returns inspection evidence; no notes part is created by these APIs.
type NotesTarget struct {
	session          *EditSession
	generation       uint64
	part, hash, text string
	doc              *losslessxml.Document
	leaf             losslessxml.Element
	consumed         bool
	paragraphs       []losslessxml.Element
	pPr, rPr, endPr  *losslessxml.NewElement
	singleLeaf       bool
}

func (t *NotesTarget) Text() string {
	if t == nil {
		return ""
	}
	return t.text
}
func notesAttr(e losslessxml.Element, local string) string {
	for _, a := range e.Attributes() {
		if a.Name == (xml.Name{Local: local}) {
			return a.Value
		}
	}
	return ""
}
func notesChildren(doc *losslessxml.Document, parent losslessxml.Element, want xml.Name) []losslessxml.Element {
	var out []losslessxml.Element
	for _, e := range doc.Elements() {
		if p, ok := e.Parent(); ok && p == parent && e.Name() == want {
			out = append(out, e)
		}
	}
	return out
}
func notesWithin(e, parent losslessxml.Element) bool {
	for p, ok := e.Parent(); ok; p, ok = p.Parent() {
		if p == parent {
			return true
		}
	}
	return false
}

// FindNotes supports an existing ordinary body placeholder with plain paragraphs
// and runs, including empty paragraphs. Fields, text locks and extended structures
// refuse. Other placeholders are retained untouched, including slide-number fields.
func (s *EditSession) FindNotes(slidePart string) (*NotesTarget, error) {
	if !s.slides[slidePart] {
		return nil, editRefusal("missing_target", "notes source is not an enrolled slide")
	}
	g, err := s.pkg.Graph()
	if err != nil {
		return nil, err
	}
	parts := map[string]packaging.GraphPart{}
	for _, p := range g.Parts {
		parts[p.Name] = p
	}
	part := ""
	for _, edge := range g.Edges {
		if edge.Source != slidePart || edge.Type != packaging.RelTypeNotesSlide {
			continue
		}
		if part != "" || edge.External || strings.Contains(edge.Target, "#") {
			return nil, editRefusal("relationship_policy", "ambiguous/external notes relationship")
		}
		part = edge.ResolvedPart
	}
	if part == "" {
		return nil, editRefusal("missing_target", "slide has no existing notes part")
	}
	if parts[part].ContentType != packaging.ContentTypeNotesSlide || parts[part].Inbound != 1 {
		return nil, editRefusal("relationship_policy", "notes part is shared or mistyped")
	}
	for _, p := range g.Parts {
		if p.ContentType != packaging.ContentTypePresentation {
			continue
		}
		b, _, err := s.pkg.Part(p.Name)
		if err != nil {
			return nil, err
		}
		d, err := losslessxml.Parse(b)
		if err != nil {
			return nil, err
		}
		for _, e := range d.Elements() {
			if e.Name() == name(packaging.NSPresentationML, "modifyVerifier") {
				return nil, editRefusal("protected_operation", "presentation modification protection")
			}
		}
	}
	b, hash, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	doc, err := losslessxml.Parse(b)
	if err != nil {
		return nil, editRefusal("unsupported_structure", err.Error())
	}
	es := doc.Elements()
	root := es[0]
	if root.Name() != name(packaging.NSPresentationML, "notes") {
		return nil, editRefusal("unsupported_structure", "notes root required")
	}
	var shape losslessxml.Element
	count := 0
	for _, e := range es {
		n := e.Name()
		if n.Space != packaging.NSPresentationML && n.Space != packaging.NSDrawingML {
			return nil, editRefusal("unsupported_structure", "unknown notes namespace")
		}
		if n.Local == "extLst" {
			return nil, editRefusal("unsupported_structure", "notes extension unsupported")
		}
		if n != name(packaging.NSPresentationML, "ph") || notesAttr(e, "type") != "body" {
			continue
		}
		count++
		p := e
		for _, local := range []string{"nvPr", "nvSpPr", "sp", "spTree", "cSld", "notes"} {
			next, ok := p.Parent()
			if !ok || next.Name() != name(packaging.NSPresentationML, local) {
				return nil, editRefusal("unsupported_structure", "unexpected notes body owner")
			}
			if local == "sp" {
				shape = next
			}
			p = next
		}
		if p != root {
			return nil, editRefusal("unsupported_structure", "nested notes root")
		}
	}
	if count == 0 {
		return nil, editRefusal("missing_target", "notes body absent")
	}
	if count != 1 {
		return nil, editRefusal("ambiguous_target", "duplicate notes body placeholders")
	}
	bodies := notesChildren(doc, shape, name(packaging.NSPresentationML, "txBody"))
	if len(bodies) != 1 {
		return nil, editRefusal("unsupported_structure", "one notes text body required")
	}
	body := bodies[0]
	for _, e := range es {
		if !notesWithin(e, shape) && e != shape {
			continue
		}
		n := e.Name()
		if n == name(packaging.NSDrawingML, "spLocks") {
			value, leaf := e.Text()
			if !leaf || strings.TrimSpace(value) != "" {
				return nil, editRefusal("unsupported_structure", "extended notes lock structure")
			}
			// Grouping is not performed by this text-only operation. Unknown
			// locks (including text-edit locks) keep the conservative refusal.
			for _, a := range e.Attributes() {
				if a.Name != (xml.Name{Local: "noGrp"}) {
					return nil, editRefusal("protected_operation", "notes shape has an unproved lock")
				}
				if a.Value != "0" && a.Value != "1" && a.Value != "true" && a.Value != "false" {
					return nil, editRefusal("unsupported_structure", "invalid notes grouping lock")
				}
			}
		}
		for _, a := range e.Attributes() {
			if a.Name.Space != "" && a.Name.Space != "http://www.w3.org/XML/1998/namespace" {
				return nil, editRefusal("unsupported_structure", "unknown notes body attribute namespace")
			}
		}
	}
	target := &NotesTarget{session: s, generation: s.generation, part: part, hash: hash, doc: doc}
	if err = readNotesParagraphs(doc, body, target); err != nil {
		return nil, err
	}
	return target, nil
}

// ReplaceNotes uses the first paragraph/run formatting for newline-separated
// paragraphs, retaining body metadata and other placeholders. It never creates
// a notes graph or invokes the legacy notes_slide accessor. Changed targets stale
// all held session handles, while identical text keeps its target reusable.
func (s *EditSession) ReplaceNotes(target *NotesTarget, text string) error {
	if target == nil || target.session != s || target.generation != s.generation || target.consumed {
		return editRefusal("stale_target", "stale, consumed or foreign notes target")
	}
	_, hash, err := s.pkg.Part(target.part)
	if err != nil {
		return err
	}
	if hash != target.hash {
		return editRefusal("stale_target", "notes source fingerprint changed")
	}
	if strings.ContainsAny(text, "\r\t") {
		return editRefusal("unsupported_structure", "notes tabs/carriage returns require explicit structure")
	}
	if text == target.text {
		return nil
	}
	var b []byte
	if target.singleLeaf && text != "" && !strings.Contains(text, "\n") {
		b, err = target.doc.ReplaceText([]losslessxml.TextEdit{{Target: target.leaf, Text: text}})
	} else {
		nodes := []losslessxml.NewElement{}
		for _, line := range strings.Split(text, "\n") {
			nodes = append(nodes, notesNewParagraph(target, line))
		}
		edits := []losslessxml.ElementReplacement{{Target: target.paragraphs[0], Nodes: nodes}}
		for _, p := range target.paragraphs[1:] {
			edits = append(edits, losslessxml.ElementReplacement{Target: p})
		}
		b, err = target.doc.ReplaceElements(edits)
	}
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
