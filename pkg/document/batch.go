package document

import (
	"sort"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// TextReplacement selects a previously inspected target and its replacement.
type TextReplacement struct {
	Target *TextTarget
	Text   string
}
type leafChange struct {
	lo, hi int
	text   string
}

// ReplaceBatch commits all supplied targets together. Any refusal aborts it,
// unlike extended's replace_all which records independent per-match refusals.
// Overlapping selections refuse even if their requested values are identical.
func (s *EditSession) ReplaceBatch(changes []TextReplacement) error {
	if len(changes) == 0 {
		return nil
	}
	source, hash, err := s.pkg.Part(s.part)
	if err != nil {
		return err
	}
	d, err := losslessxml.Parse(source)
	if err != nil {
		return err
	}
	elements := d.Elements()
	type selected struct{ paragraph, start, end int }
	var selections []selected
	staged := map[int][]leafChange{}
	changed := []*TextTarget{}
	seen := map[*TextTarget]bool{}
	for _, change := range changes {
		target := change.Target
		if target == nil || target.session != s || target.consumed || target.generation != s.generation || target.hash != hash {
			return editRefusal("stale_target", "stale, foreign or consumed batch target")
		}
		if seen[target] {
			return editRefusal("ambiguous_target", "duplicate selected target")
		}
		seen[target] = true
		selection := selected{target.paragraph.Ordinal(), target.start, target.end}
		for _, prior := range selections {
			if prior.paragraph == selection.paragraph && max(prior.start, selection.start) < min(prior.end, selection.end) {
				return editRefusal("ambiguous_target", "overlapping selected spans")
			}
		}
		selections = append(selections, selection)
		if change.Text == target.text {
			continue
		}
		edits, err := s.planText(target, change.Text)
		if err != nil {
			return err
		}
		changed = append(changed, target)
		// Convert the proved old/new leaf edits to rune deltas, then merge deltas on
		// the original immutable leaf. This avoids last-writer-wins in shared leaves.
		for _, edit := range edits {
			old, leaf := edit.Target.Text()
			if !leaf {
				return editRefusal("unsupported_structure", "non-leaf batch edit")
			}
			if old == edit.Text {
				continue
			}
			lo, hi, newEnd, err := changedInterval([]rune(old), []rune(edit.Text))
			if err != nil {
				return err
			}
			replacement := string([]rune(edit.Text)[lo:newEnd])
			index := edit.Target.Ordinal()
			if index < 0 || index >= len(elements) {
				return editRefusal("stale_target", "invalid element identity")
			}
			canonical, ok := elements[index].Text()
			if !ok || canonical != old {
				return editRefusal("stale_target", "snapshot leaf differs")
			}
			for _, prior := range staged[index] {
				if max(prior.lo, lo) < min(prior.hi, hi) || prior.lo == lo && (prior.lo == prior.hi || lo == hi) {
					return editRefusal("ambiguous_target", "overlapping leaf deltas")
				}
			}
			staged[index] = append(staged[index], leafChange{lo, hi, replacement})
		}
	}
	if len(changed) == 0 {
		return nil
	}
	edits := []losslessxml.TextEdit{}
	for index, deltas := range staged {
		text, _ := elements[index].Text()
		runes := []rune(text)
		sort.Slice(deltas, func(i, j int) bool { return deltas[i].lo < deltas[j].lo })
		result := []rune{}
		cursor := 0
		for _, delta := range deltas {
			if delta.lo < cursor {
				return editRefusal("ambiguous_target", "overlapping batch offsets")
			}
			result = append(result, runes[cursor:delta.lo]...)
			result = append(result, []rune(delta.text)...)
			cursor = delta.hi
		}
		result = append(result, runes[cursor:]...)
		probe := *changes[0].Target
		probe.doc = d
		probe.element = elements[index]
		if err = s.guard(&probe, string(result)); err != nil {
			return err
		}
		edits = append(edits, losslessxml.TextEdit{Target: elements[index], Text: string(result)})
	}
	data, err := d.Edit(edits, whitespaceAttributes(edits))
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: s.part, ExpectedSHA256: hash, Data: data}}); err != nil {
		return err
	}
	s.generation++
	for _, target := range changed {
		target.consumed = true
	}
	return nil
}
