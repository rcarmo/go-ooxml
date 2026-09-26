package document

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type View string

const (
	CurrentView  View = "current"
	OriginalView View = "original"
	AllView      View = "all"
)

// StoryBlock is inspection evidence only; it cannot authorise a mutation.
// Paragraph is a zero-based index in source order within its story part.
type StoryBlock struct {
	Part      string `json:"part"`
	Paragraph int    `json:"paragraph"`
	Text      string `json:"text"`
	InTextBox bool   `json:"in_text_box"`
}
type StoryOutline struct {
	Schema     int          `json:"schema"`
	View       View         `json:"view"`
	Blocks     []StoryBlock `json:"blocks"`
	Unreadable []string     `json:"unreadable"`
}

// Outline projects paragraph text through basic insertion/deletion/move wrappers.
// It reports unsupported alternate content and field instructions instead of
// pretending their application-dependent display has been evaluated.
func (s *EditSession) Outline(view View) (StoryOutline, error) {
	if view != CurrentView && view != OriginalView && view != AllView {
		return StoryOutline{}, fmt.Errorf("invalid story view %q", view)
	}
	g, err := s.pkg.Graph()
	if err != nil {
		return StoryOutline{}, err
	}
	storyTypes := map[string]bool{packaging.RelTypeHeader: true, packaging.RelTypeFooter: true, packaging.RelTypeFootnotes: true, packaging.RelTypeEndnotes: true, packaging.RelTypeComments: true}
	parts := []string{s.part}
	seen := map[string]bool{s.part: true}
	out := StoryOutline{Schema: 1, View: view, Blocks: []StoryBlock{}, Unreadable: []string{}}
	for i := 0; i < len(parts); i++ {
		for _, e := range g.Edges {
			if e.Source != parts[i] || !storyTypes[e.Type] {
				continue
			}
			if e.External {
				out.Unreadable = append(out.Unreadable, e.Source+":"+e.ID+": external story")
				continue
			}
			if !seen[e.ResolvedPart] {
				seen[e.ResolvedPart] = true
				parts = append(parts, e.ResolvedPart)
			}
		}
	}
	sort.Strings(parts[1:])
	for _, part := range parts {
		b, _, err := s.pkg.Part(part)
		if err != nil {
			return StoryOutline{}, err
		}
		d, err := losslessxml.Parse(b)
		if err != nil {
			out.Unreadable = append(out.Unreadable, part+": "+err.Error())
			continue
		}
		blocks, warnings := outlinePart(d, part, view)
		out.Blocks = append(out.Blocks, blocks...)
		out.Unreadable = append(out.Unreadable, warnings...)
	}
	return out, nil
}
func visible(e losslessxml.Element, view View) bool {
	for p := e; ; {
		n := p.Name()
		if n.Space == packaging.NSWordprocessingML {
			if view == CurrentView && (n.Local == "del" || n.Local == "moveFrom") {
				return false
			}
			if view == OriginalView && (n.Local == "ins" || n.Local == "moveTo") {
				return false
			}
		}
		parent, ok := p.Parent()
		if !ok {
			break
		}
		p = parent
	}
	return true
}
func ancestor(e losslessxml.Element, ns, local string) (losslessxml.Element, bool) {
	for p, ok := e.Parent(); ok; p, ok = p.Parent() {
		if p.Name() == (xml.Name{Space: ns, Local: local}) {
			return p, true
		}
	}
	return losslessxml.Element{}, false
}
func outlinePart(d *losslessxml.Document, part string, view View) ([]StoryBlock, []string) {
	blocks := []StoryBlock{}
	warnings := []string{}
	seenWarnings := map[string]bool{}
	warn := func(what string) {
		if !seenWarnings[what] {
			warnings = append(warnings, part+": "+what)
			seenWarnings[what] = true
		}
	}
	es := d.Elements()
	indices := map[losslessxml.Element]int{}
	texts := map[losslessxml.Element]*strings.Builder{}
	ordinal := 0
	for _, e := range es {
		n := e.Name()
		if n.Space == "http://schemas.openxmlformats.org/markup-compatibility/2006" && n.Local == "AlternateContent" {
			warn("alternate content requires branch selection")
		}
		if n.Space != packaging.NSWordprocessingML {
			continue
		}
		switch n.Local {
		case "instrText", "fldSimple":
			warn("field instructions are not evaluated")
		case "altChunk":
			warn("alternate chunk content is not expanded")
		case "pPrChange", "rPrChange", "tblPrChange", "trPrChange", "tcPrChange":
			warn("property revisions are not projected")
		}
		if n.Local != "p" {
			continue
		}
		current := ordinal
		ordinal++
		if !visible(e, view) {
			continue
		}
		if _, ok := ancestor(e, "http://schemas.openxmlformats.org/markup-compatibility/2006", "AlternateContent"); ok {
			continue
		}
		_, box := ancestor(e, packaging.NSWordprocessingML, "txbxContent")
		indices[e] = len(blocks)
		texts[e] = &strings.Builder{}
		blocks = append(blocks, StoryBlock{Part: part, Paragraph: current, InTextBox: box})
	}
	for _, e := range es {
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
		value := ""
		switch n.Local {
		case "t", "delText":
			if n.Local == "delText" && view == CurrentView {
				continue
			}
			text, leaf := e.Text()
			if !leaf {
				warn("mixed text element skipped")
				continue
			}
			value = text
		case "tab":
			value = "\t"
		case "br", "cr":
			value = "\n"
		case "noBreakHyphen":
			value = "\u2011"
		case "softHyphen":
			value = "\u00ad"
		default:
			continue
		}
		owner, ok := ancestor(e, packaging.NSWordprocessingML, "p")
		if !ok {
			continue
		}
		if b, ok := texts[owner]; ok {
			b.WriteString(value)
		}
	}
	for e, index := range indices {
		blocks[index].Text = texts[e].String()
	}
	return blocks, warnings
}
