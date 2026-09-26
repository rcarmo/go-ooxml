package document

import (
	"encoding/xml"
	"strings"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Story is the exact package part containing the selected text.
func (t *TextTarget) Story() string {
	if t == nil {
		return ""
	}
	return t.story
}

type textSegment struct {
	element    losslessxml.Element
	text       string
	start, end int
}
type textAtom struct {
	element    losslessxml.Element
	text       string
	start, end int
}
type paragraphText struct {
	paragraph losslessxml.Element
	atoms     []textAtom
	text      []rune
}

// FindText returns every non-overlapping exact match, in source order, within
// current-view paragraphs of the main story. TextEvidence retains original
// Unicode codepoints; run boundaries insert no characters. Cross-paragraph and
// normalised search are not implied by this method.
func (s *EditSession) FindText(text string) ([]*TextTarget, error) {
	if !utf8.ValidString(text) {
		return nil, editRefusal("unsupported_structure", "search text is not UTF-8")
	}
	if text == "" {
		return []*TextTarget{}, nil
	}
	data, hash, err := s.pkg.Part(s.part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	return s.findExact(d, hash, text), nil
}

// Text is immutable inspection evidence for the selected source span.
func (t *TextTarget) Text() string {
	if t == nil {
		return ""
	}
	return t.text
}

func paragraphStreams(d *losslessxml.Document) []paragraphText {
	return paragraphStreamsView(d, CurrentView)
}
func paragraphStreamsView(d *losslessxml.Document, view View) []paragraphText {
	streams := []paragraphText{}
	indices := map[losslessxml.Element]int{}
	for _, e := range d.Elements() {
		if e.Name() == (xml.Name{Space: packaging.NSWordprocessingML, Local: "p"}) && visible(e, view) {
			indices[e] = len(streams)
			streams = append(streams, paragraphText{paragraph: e})
		}
	}
	for _, e := range d.Elements() {
		if !visible(e, view) {
			continue
		}
		if _, ok := ancestor(e, "http://schemas.openxmlformats.org/markup-compatibility/2006", "AlternateContent"); ok {
			continue
		}
		n := e.Name()
		if n.Space != packaging.NSWordprocessingML {
			continue
		}
		text := ""
		switch n.Local {
		case "t", "delText":
			if n.Local == "delText" && view == CurrentView {
				continue
			}
			got, leaf := e.Text()
			if !leaf {
				continue
			}
			text = got
		case "tab":
			text = "\t"
		case "br", "cr":
			text = "\n"
		case "noBreakHyphen":
			text = "\u2011"
		case "softHyphen":
			text = "\u00ad"
		default:
			continue
		}
		owner, ok := ancestor(e, packaging.NSWordprocessingML, "p")
		if !ok {
			continue
		}
		i, ok := indices[owner]
		if !ok {
			continue
		}
		s := &streams[i]
		start := len(s.text)
		s.text = append(s.text, []rune(text)...)
		s.atoms = append(s.atoms, textAtom{e, text, start, len(s.text)})
	}
	return streams
}
func (s *EditSession) findExact(d *losslessxml.Document, hash, text string) []*TextTarget {
	out := []*TextTarget{}
	if text == "" || !utf8.ValidString(text) {
		return out
	}
	needle := []rune(text)
	for _, p := range paragraphStreams(d) {
		raw := string(p.text)
		byteOffset := 0
		for byteOffset <= len(raw) {
			relative := strings.Index(raw[byteOffset:], text)
			if relative < 0 {
				break
			}
			at := byteOffset + relative
			start := utf8.RuneCountInString(raw[:at])
			end := start + len(needle)
			segments := []textSegment{}
			for _, a := range p.atoms {
				if a.end <= start || a.start >= end {
					continue
				}
				lo, hi := max(start, a.start)-a.start, min(end, a.end)-a.start
				segments = append(segments, textSegment{a.element, a.text, lo, hi})
			}
			if len(segments) > 0 {
				out = append(out, &TextTarget{session: s, generation: s.generation, doc: d, element: segments[0].element, hash: hash, text: text, story: s.part, view: CurrentView, segments: segments, paragraph: p.paragraph, start: start, end: end})
			}
			byteOffset = at + len(text)
		}
	}
	return out
}
