package document

import (
	"bytes"
	"encoding/xml"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"reflect"
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
			if inserted {
				position += count
				continue
			}
			if start == position+count && i < len(target.segments)-1 {
				if !sameRunFormat(target.doc, seg.element, target.segments[i+1].element) {
					return nil, editRefusal("unsupported_structure", "run boundary formatting or ownership differs")
				}
				probe := *target
				probe.element = target.segments[i+1].element
				if err = s.guard(&probe, target.segments[i+1].text); err != nil {
					return nil, err
				}
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

// sameRunFormat is conservative: raw property markup and expanded run/text
// attributes must agree. Equivalent differently serialised properties may refuse.
func sameRunFormat(d *losslessxml.Document, a, b losslessxml.Element) bool {
	ra, oka := a.Parent()
	rb, okb := b.Parent()
	if !oka || !okb || ra.Name() != (xml.Name{Space: packaging.NSWordprocessingML, Local: "r"}) || rb.Name() != ra.Name() {
		return false
	}
	pa, oka := ra.Parent()
	pb, okb := rb.Parent()
	if !oka || !okb || pa != pb {
		return false
	}
	if !reflect.DeepEqual(ra.Attributes(), rb.Attributes()) || !reflect.DeepEqual(a.Attributes(), b.Attributes()) {
		return false
	}
	property := func(run losslessxml.Element) ([]byte, bool) {
		var raw []byte
		count := 0
		for _, e := range d.Elements() {
			p, ok := e.Parent()
			if ok && p == run && e.Name() == (xml.Name{Space: packaging.NSWordprocessingML, Local: "rPr"}) {
				raw = e.Raw()
				count++
			}
		}
		return raw, count <= 1
	}
	x, okx := property(ra)
	y, oky := property(rb)
	return okx && oky && bytes.Equal(x, y)
}
