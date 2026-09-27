package losslessxml

import (
	"bytes"
	"fmt"
	"sort"
)

// ElementReplacement replaces a non-root element with zero or more structured
// nodes rendered in its parent's namespace context. No raw XML is accepted.
type ElementReplacement struct {
	Target Element
	Nodes  []NewElement
}

// ReplaceElements preflights complete, disjoint subtrees and returns new bytes.
// Namespace declarations on a removed target do not leak into its replacement.
// Root, foreign, duplicate and overlapping selections refuse atomically.
func (d *Document) ReplaceElements(edits []ElementReplacement) ([]byte, error) {
	ordered := append([]ElementReplacement(nil), edits...)
	for _, edit := range ordered {
		e := edit.Target
		if e.doc != d || !e.valid() {
			return nil, fmt.Errorf("foreign subtree replacement")
		}
		if d.nodes[e.index].parent < 0 {
			return nil, fmt.Errorf("root subtree replacement unsupported")
		}
	}
	sort.Slice(ordered, func(i, j int) bool {
		return d.nodes[ordered[i].Target.index].start < d.nodes[ordered[j].Target.index].start
	})
	var out bytes.Buffer
	cursor := 0
	for _, edit := range ordered {
		n := d.nodes[edit.Target.index]
		if n.start < cursor {
			return nil, fmt.Errorf("overlapping subtree replacements")
		}
		parent := d.nodes[n.parent]
		var replacement bytes.Buffer
		for _, node := range edit.Nodes {
			if err := renderElement(&replacement, node, parent.ns, 0); err != nil {
				return nil, err
			}
		}
		out.Write(d.source[cursor:n.start])
		out.Write(replacement.Bytes())
		cursor = n.end
	}
	out.Write(d.source[cursor:])
	if _, err := Parse(out.Bytes()); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}
