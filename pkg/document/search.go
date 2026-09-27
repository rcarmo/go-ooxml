package document

import (
	"fmt"
	"sort"
	"strings"
	"unicode"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// SearchOptions explicitly selects normalisation or context ranking. Nth is
// one-based; zero keeps all candidates. Nth and Near cannot be combined.
type SearchOptions struct {
	Normalized bool
	Near       string
	Nth        int
	Story      string
	View       View
}

func (s *EditSession) Search(text string, options SearchOptions) ([]*TextTarget, error) {
	if options.Nth < 0 || options.Nth > 0 && options.Near != "" {
		return nil, fmt.Errorf("Nth must be nonnegative and cannot combine with Near")
	}
	if !utf8.ValidString(text) || !utf8.ValidString(options.Near) {
		return nil, fmt.Errorf("search strings must be UTF-8")
	}
	view := options.View
	if view == "" {
		view = CurrentView
	}
	if view != CurrentView && view != OriginalView && view != AllView {
		return nil, fmt.Errorf("invalid search view %q", view)
	}
	story := options.Story
	if story == "" {
		story = s.part
	}
	if story != s.part {
		parts, err := s.storyParts()
		if err != nil {
			return nil, err
		}
		found := false
		for _, part := range parts {
			if part == story {
				found = true
			}
		}
		if !found {
			return nil, editRefusal("missing_target", "requested story is not related to document")
		}
	}
	data, hash, err := s.pkg.Part(story)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	needle, _ := matchSpace([]rune(text), options.Normalized)
	if len(needle) == 0 || options.Normalized && strings.TrimSpace(string(needle)) == "" {
		return []*TextTarget{}, nil
	}
	type ranked struct {
		target   *TextTarget
		position int
	}
	var candidates []ranked
	var contexts []int
	contextNeedle, _ := matchSpace([]rune(options.Near), options.Normalized)
	paragraphs := paragraphStreamsView(d, view)
	raw := []rune{}
	atoms := []textAtom{}
	owners := []losslessxml.Element{}
	for _, p := range paragraphs {
		if len(p.atoms) == 0 {
			continue
		}
		if len(raw) > 0 {
			raw = append(raw, '\n')
		}
		offset := len(raw)
		raw = append(raw, p.text...)
		for _, a := range p.atoms {
			a.start += offset
			a.end += offset
			atoms = append(atoms, a)
			owners = append(owners, p.paragraph)
		}
	}
	value, mapping := matchSpace(raw, options.Normalized)
	for _, interval := range matchIntervals(value, needle, mapping) {
		start, end := interval[0], interval[1]
		segments := []textSegment{}
		var owner losslessxml.Element
		cross := false
		ownerOffset := 0
		for i, a := range atoms {
			if a.end <= start || a.start >= end {
				continue
			}
			if len(segments) == 0 {
				owner = owners[i]
				for j := 0; j <= i; j++ {
					if owners[j] == owner {
						ownerOffset = atoms[j].start
						break
					}
				}
			} else if owners[i] != owner {
				cross = true
			}
			segments = append(segments, textSegment{a.element, a.text, max(start, a.start) - a.start, min(end, a.end) - a.start})
		}
		if len(segments) == 0 {
			continue
		}
		covered := 0
		for _, seg := range segments {
			covered += seg.end - seg.start
		}
		if covered != end-start {
			cross = true
		}
		if first := segments[0]; first.start == first.end {
			continue
		}
		target := &TextTarget{session: s, generation: s.generation, doc: d, element: segments[0].element, hash: hash, text: string(raw[start:end]), segments: segments, paragraph: owner, start: start - ownerOffset, end: end - ownerOffset, story: story, view: view, crossParagraph: cross}
		candidates = append(candidates, ranked{target, start})
	}
	if len(contextNeedle) > 0 {
		for _, interval := range matchIntervals(value, contextNeedle, mapping) {
			contexts = append(contexts, interval[0])
		}
	}
	if options.Near != "" && len(contexts) > 0 {
		distance := func(position int) int {
			best := int(^uint(0) >> 1)
			for _, at := range contexts {
				d := position - at
				if d < 0 {
					d = -d
				}
				if d < best {
					best = d
				}
			}
			return best
		}
		sort.SliceStable(candidates, func(i, j int) bool { return distance(candidates[i].position) < distance(candidates[j].position) })
	}
	if options.Nth > 0 {
		if options.Nth > len(candidates) {
			return nil, editRefusal("missing_target", "requested occurrence absent")
		}
		return []*TextTarget{candidates[options.Nth-1].target}, nil
	}
	out := make([]*TextTarget, len(candidates))
	for i, c := range candidates {
		out[i] = c.target
	}
	return out, nil
}
func matchSpace(raw []rune, normalised bool) ([]rune, []int) {
	out := []rune{}
	mapping := []int{}
	for i, r := range raw {
		if !normalised {
			out = append(out, r)
			mapping = append(mapping, i)
			continue
		}
		switch r {
		case '\u00ad':
			continue
		case '‘', '’':
			r = '\''
		case '“', '”':
			r = '"'
		case '–', '—', '−':
			r = '-'
		}
		if unicode.IsSpace(r) || r >= 0x1c && r <= 0x1f {
			r = ' '
		}
		folded, ok := fullCaseFold[r]
		if !ok {
			folded = string(r)
		}
		for _, f := range folded {
			if f == ' ' && len(out) > 0 && out[len(out)-1] == ' ' {
				continue
			}
			out = append(out, f)
			mapping = append(mapping, i)
		}
	}
	return out, mapping
}
func matchIntervals(value, needle []rune, mapping []int) [][2]int {
	intervals := [][2]int{}
	if len(needle) == 0 {
		return intervals
	}
	for i := 0; i+len(needle) <= len(value); {
		equal := true
		for j, r := range needle {
			if value[i+j] != r {
				equal = false
				break
			}
		}
		if !equal {
			i++
			continue
		}
		end := i + len(needle)
		if (i == 0 || mapping[i-1] != mapping[i]) && (end == len(mapping) || mapping[end-1] != mapping[end]) {
			intervals = append(intervals, [2]int{mapping[i], mapping[end-1] + 1})
			i = end
		} else {
			i++
		}
	}
	return intervals
}
