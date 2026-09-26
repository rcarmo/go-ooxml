package document

import (
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// changedInterval considers all maximal non-overlapping prefix/suffix splits.
// Repeated affixes can yield several equally good intervals; never choose one.
func changedInterval(old, new []rune) (start, end, newEnd int, err error) {
	prefix := 0
	for prefix < len(old) && prefix < len(new) && old[prefix] == new[prefix] {
		prefix++
	}
	suffix := 0
	for suffix < len(old) && suffix < len(new) && old[len(old)-1-suffix] == new[len(new)-1-suffix] {
		suffix++
	}
	total := min(min(len(old), len(new)), prefix+suffix)
	low, high := max(0, total-suffix), min(prefix, total)
	if low != high {
		return 0, 0, 0, editRefusal("ambiguous_target", "repeated exact affixes identify multiple changed intervals")
	}
	suffix = total - low
	return low, len(old) - suffix, len(new) - suffix, nil
}

func (s *EditSession) planText(target *TextTarget, replacement string) ([]losslessxml.TextEdit, error) {
	if target.view != "" && target.view != CurrentView {
		return nil, editRefusal("unsupported_structure", "historical search targets are inspection-only")
	}
	if target.crossParagraph {
		return nil, editRefusal("boundary_violation", "span crosses paragraph ownership")
	}
	if target.story != "" && target.story != s.part {
		return nil, editRefusal("unsupported_structure", "related-story editing is not implemented")
	}
	old, next := []rune(target.text), []rune(replacement)
	start, end, newEnd, err := changedInterval(old, next)
	if err != nil {
		return nil, err
	}
	residual := string(next[start:newEnd])
	edits := []losslessxml.TextEdit{}
	position := 0
	inserted := false
	for i, seg := range target.segments {
		count := seg.end - seg.start
		lo, hi := max(start, position), min(end, position+count)
		insertion := start == end && start >= position && start <= position+count
		if insertion {
			// Pure insertion at an interior run boundary needs full run/owner equality.
			// Until that proof is implemented, refuse rather than picking an arbitrary run.
			if start == position && i > 0 || start == position+count && i < len(target.segments)-1 {
				return nil, editRefusal("unsupported_structure", "insertion at run boundary needs formatting/owner proof")
			}
			lo = start
			hi = start
		}
		if lo < hi || insertion {
			if seg.element.Name().Space != packaging.NSWordprocessingML || seg.element.Name().Local != "t" {
				return nil, editRefusal("unsupported_structure", "changed interval contains non-text atom")
			}
			text := []rune(seg.text)
			a, b := seg.start+lo-position, seg.start+hi-position
			fragment := ""
			if !inserted {
				fragment = residual
				inserted = true
			}
			result := string(text[:a]) + fragment + string(text[b:])
			probe := *target
			probe.element = seg.element
			if err = s.guard(&probe, result); err != nil {
				return nil, err
			}
			edits = append(edits, losslessxml.TextEdit{Target: seg.element, Text: result})
		}
		position += count
	}
	if !inserted {
		return nil, editRefusal("unsupported_structure", "no concrete changed text owner")
	}
	return edits, nil
}
