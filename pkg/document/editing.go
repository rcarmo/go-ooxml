package document

import (
	"encoding/xml"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// EditSession is a preservation-first editor for a deliberately bounded subset:
// exact complete plain w:t leaves in direct body paragraphs. It does not replace
// the legacy Document interfaces or claim cross-run/story/review parity.
// Sessions are single-owner; use separate sessions for concurrent editing.
type EditSession struct {
	pkg        *packaging.Preserved
	part       string
	generation uint64
}

// TextTarget authorises a single snapshot-bound leaf edit; fields are private.
type TextTarget struct {
	session    *EditSession
	generation uint64
	doc        *losslessxml.Document
	element    losslessxml.Element
	hash       string
	text       string
	consumed   bool
	segments   []textSegment
	paragraph  losslessxml.Element
	start, end int
}

func editRefusal(kind, detail string) error {
	return &packaging.Refusal{Kind: kind, Operation: "word_edit", Detail: detail}
}

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
				return nil, editRefusal("ambiguous_target", "multiple main document relationships")
			}
			main = e.ResolvedPart
		}
	}
	if main == "" {
		return nil, editRefusal("missing_target", "Word main part not found")
	}
	for _, part := range g.Parts {
		if part.Name == main && part.ContentType != packaging.ContentTypeWordDocument {
			return nil, editRefusal("unsupported_structure", "only ordinary transitional DOCX is supported")
		}
	}
	data, _, err := p.Part(main)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	if d.Elements()[0].Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: "document"}) {
		return nil, editRefusal("unsupported_structure", "not a transitional Word document")
	}
	return &EditSession{pkg: p, part: main}, nil
}

func (s *EditSession) FindOne(text string) (*TextTarget, error) {
	if text == "" {
		return nil, editRefusal("ambiguous_target", "empty text is not a target")
	}
	data, hash, err := s.pkg.Part(s.part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	matches := s.findExact(d, hash, text)
	if len(matches) == 0 {
		return nil, editRefusal("missing_target", "no exact text span matched")
	}
	if len(matches) != 1 {
		return nil, editRefusal("ambiguous_target", "multiple exact text spans")
	}
	return matches[0], nil
}

func (s *EditSession) guard(target *TextTarget, replacement string) error {
	g, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	for _, edge := range g.Edges {
		if edge.Source == s.part && edge.Type == packaging.RelTypeSettings {
			if edge.External {
				return editRefusal("unsupported_structure", "external settings")
			}
			b, _, err := s.pkg.Part(edge.ResolvedPart)
			if err != nil {
				return err
			}
			settings, err := losslessxml.Parse(b)
			if err != nil {
				return err
			}
			if settings.Elements()[0].Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: "settings"}) {
				return editRefusal("unsupported_structure", "unknown settings root")
			}
			for _, e := range settings.Elements() {
				if e.Name().Space != packaging.NSWordprocessingML {
					return editRefusal("unsupported_structure", "unknown settings extension")
				}
				for _, a := range e.Attributes() {
					if a.Name.Space != "" && a.Name.Space != packaging.NSWordprocessingML && a.Name.Space != "http://www.w3.org/XML/1998/namespace" {
						return editRefusal("unsupported_structure", "unknown settings attribute namespace")
					}
				}
				if e.Name() == (xml.Name{Space: packaging.NSWordprocessingML, Local: "documentProtection"}) {
					enforcement := ""
					for _, a := range e.Attributes() {
						if a.Name == (xml.Name{Space: packaging.NSWordprocessingML, Local: "enforcement"}) {
							enforcement = a.Value
						}
					}
					switch enforcement {
					case "0", "false", "off":
					case "1", "true", "on":
						return editRefusal("protected_operation", "active document protection")
					default:
						return editRefusal("unsupported_structure", "unknown protection enforcement")
					}
				}
				if e.Name() == (xml.Name{Space: packaging.NSWordprocessingML, Local: "trackRevisions"}) {
					return editRefusal("unsupported_structure", "track-revisions settings require review-aware editing")
				}
			}
		}
	}
	blocked := map[string]bool{"ins": true, "del": true, "moveFrom": true, "moveTo": true, "fldChar": true, "fldSimple": true, "instrText": true, "commentRangeStart": true, "commentRangeEnd": true, "bookmarkStart": true, "bookmarkEnd": true, "permStart": true, "permEnd": true}
	for _, e := range target.doc.Elements() {
		if e.Name().Space == packaging.NSWordprocessingML && (blocked[e.Name().Local] || strings.HasSuffix(e.Name().Local, "Change")) {
			return editRefusal("unsupported_structure", "review/field/range markup requires a specialised editor")
		}
	}
	// Restrict ownership to t/r/p/body/document; tables, content controls, hyperlinks,
	// text boxes and other stories are explicitly outside this first subset.
	parent := target.element
	for _, name := range []string{"r", "p", "body", "document"} {
		var ok bool
		parent, ok = parent.Parent()
		if !ok || parent.Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: name}) {
			return editRefusal("unsupported_structure", "text is outside a plain body run")
		}
	}
	run, _ := target.element.Parent()
	paragraph, _ := run.Parent()
	for _, e := range target.doc.Elements() {
		p, ok := e.Parent()
		if !ok {
			continue
		}
		if p == run && e != target.element && e.Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: "rPr"}) {
			return editRefusal("unsupported_structure", "run has non-text content")
		}
		if p == paragraph && e.Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: "r"}) && e.Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: "pPr"}) {
			return editRefusal("unsupported_structure", "paragraph has wrappers or markers")
		}
	}
	space := ""
	for e := target.element; ; {
		for _, a := range e.Attributes() {
			if a.Name == (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) {
				space = a.Value
				break
			}
		}
		if space != "" {
			break
		}
		p, ok := e.Parent()
		if !ok {
			break
		}
		e = p
	}
	if space != "" && space != "default" && space != "preserve" {
		return editRefusal("unsupported_structure", "unknown xml:space policy")
	}
	if strings.ContainsAny(replacement, "\t\r\n") {
		return editRefusal("unsupported_structure", "tabs and line breaks require Word elements")
	}
	return nil
}

func (s *EditSession) Replace(target *TextTarget, text string) error {
	if target == nil || target.session != s || target.consumed || target.generation != s.generation {
		return editRefusal("stale_target", "target is stale, foreign or consumed")
	}
	_, current, err := s.pkg.Part(s.part)
	if err != nil {
		return err
	}
	if current != target.hash {
		return editRefusal("stale_target", "part fingerprint changed")
	}
	if text == target.text {
		return nil
	}
	edits, err := s.planText(target, text)
	if err != nil {
		return err
	}
	data, err := target.doc.Edit(edits, whitespaceAttributes(edits))
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: s.part, ExpectedSHA256: target.hash, Data: data}}); err != nil {
		return err
	}
	s.generation++
	target.consumed = true
	return nil
}
func (s *EditSession) SaveAs(path string) (packaging.Receipt, error) { return s.pkg.SaveAs(path) }
