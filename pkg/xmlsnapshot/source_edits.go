package xmlsnapshot

import (
	"bytes"
	"fmt"
	"sort"
)

// SourceEdit replaces one original-source UTF-16 half-open interval with an
// already-escaped XML fragment. It is never interpreted as plain text.
type SourceEdit struct {
	Start, End int
	Value      string
}

const (
	CategoryEditOverlap = "edit-overlap"
	CategoryEditUnsafe  = "edit-unsafe"
)

// EditError distinguishes overlap from an unsafe resulting XML document.
type EditError struct {
	Category string
	Cause    error
}

func (e *EditError) Error() string { return e.Cause.Error() }
func (e *EditError) Unwrap() error { return e.Cause }

// ApplyEdits atomically splices disjoint UTF-16 source ranges and validates
// the entire result. Neither source nor supplied replacement list is mutated.
func (d *Document) ApplyEdits(edits []SourceEdit) ([]byte, error) {
	if d == nil || d.doc == nil {
		return nil, &EditError{CategoryEditUnsafe, fmt.Errorf("nil XML snapshot")}
	}
	if len(edits) == 0 {
		return d.Source(), nil
	}
	type splice struct {
		start, end int
		value      string
	}
	items := make([]splice, 0, len(edits))
	byteAt := func(units int) (int, error) {
		if units < 0 {
			return 0, fmt.Errorf("negative XML source offset")
		}
		for i, n := range d.offsets {
			if n == units {
				return i, nil
			}
		}
		return 0, fmt.Errorf("offset %d splits a UTF-16 scalar or exceeds source", units)
	}
	for _, edit := range edits {
		start, err := byteAt(edit.Start)
		if err != nil {
			return nil, &EditError{CategoryEditUnsafe, err}
		}
		end, err := byteAt(edit.End)
		if err != nil {
			return nil, &EditError{CategoryEditUnsafe, err}
		}
		if end < start {
			return nil, &EditError{CategoryEditUnsafe, fmt.Errorf("reversed XML range")}
		}
		items = append(items, splice{start, end, edit.Value})
	}
	sort.SliceStable(items, func(i, j int) bool { return items[i].start < items[j].start })
	var out bytes.Buffer
	cursor := 0
	previousEnd := -1
	for i, item := range items {
		if i > 0 && (item.start < previousEnd || item.start == previousEnd && item.start == item.end && items[i-1].start == items[i-1].end) {
			return nil, &EditError{CategoryEditOverlap, fmt.Errorf("overlapping XML source edits")}
		}
		out.Write(d.source[cursor:item.start])
		out.WriteString(item.value)
		cursor = item.end
		previousEnd = item.end
	}
	out.Write(d.source[cursor:])
	// The bounded parser checks DTD, structural validity and resource ceilings on
	// the new document; a refused result cannot leak partial output.
	if _, err := ParseWithLimits(out.Bytes(), DefaultLimits()); err != nil {
		return nil, &EditError{CategoryEditUnsafe, err}
	}
	return out.Bytes(), nil
}
