package losslessxml

import (
	"bytes"
	"fmt"
	"sort"
)

// RemoveElements removes complete non-root elements from an immutable snapshot.
// Sibling whitespace/comments and every other byte remain unchanged. Ancestor,
// duplicate or foreign selections refuse atomically; root removal is unsupported.
func (d *Document) RemoveElements(targets []Element) ([]byte, error) {
	ranges := []byteSplice{}
	seen := map[int]bool{}
	for _, e := range targets {
		if e.doc != d || !e.valid() {
			return nil, fmt.Errorf("foreign removal target")
		}
		n := d.nodes[e.index]
		if n.parent < 0 {
			return nil, fmt.Errorf("cannot remove document root")
		}
		if seen[e.index] {
			return nil, fmt.Errorf("duplicate removal target")
		}
		seen[e.index] = true
		ranges = append(ranges, byteSplice{start: n.start, end: n.end})
	}
	sort.Slice(ranges, func(i, j int) bool { return ranges[i].start < ranges[j].start })
	var out bytes.Buffer
	cursor := 0
	for _, r := range ranges {
		if r.start < cursor {
			return nil, fmt.Errorf("overlapping removal targets")
		}
		out.Write(d.source[cursor:r.start])
		cursor = r.end
	}
	out.Write(d.source[cursor:])
	if _, err := Parse(out.Bytes()); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}
