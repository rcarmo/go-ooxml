package acceptance

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func formattingOpeningRange(source []byte, el losslessxml.Element) (int, int, error) {
	a, b := el.SourceRange()
	if a < 0 || b < a || b > len(source) {
		return 0, 0, fmt.Errorf("invalid XML element span")
	}
	inside := source[a:b]
	quote := byte(0)
	for i, c := range inside {
		if quote != 0 {
			if c == quote {
				quote = 0
			}
			continue
		}
		if c == '\'' || c == '"' {
			quote = c
			continue
		}
		if c == '>' {
			return a, a + i + 1, nil
		}
	}
	return 0, 0, fmt.Errorf("unterminated XML opening")
}
func formattingSelected(d *losslessxml.Document, kind string) (losslessxml.Element, error) {
	if len(d.Elements()) == 0 {
		return losslessxml.Element{}, fmt.Errorf("empty XML")
	}
	root := d.Elements()[0]
	if kind == "hide" || kind == "unhide" || strings.HasPrefix(kind, "visibility-") {
		return root, nil
	}
	shape, e := pptxShape(d, 4)
	if e != nil {
		return losslessxml.Element{}, e
	}
	if kind == "move" || kind == "resize" || kind == "transform" || kind == "geometry-refusal" {
		sp := pptxChild(d, shape, pptxElem(pptxP, "spPr"))
		if len(sp) != 1 {
			return losslessxml.Element{}, fmt.Errorf("shape properties")
		}
		xs := pptxChild(d, sp[0], pptxElem(pptxA, "xfrm"))
		if len(xs) != 1 {
			return losslessxml.Element{}, fmt.Errorf("shape transform")
		}
		return xs[0], nil
	}
	body := pptxChild(d, shape, pptxElem(pptxP, "txBody"))
	if len(body) != 1 {
		return losslessxml.Element{}, fmt.Errorf("shape text body")
	}
	ps := pptxChild(d, body[0], pptxElem(pptxA, "p"))
	if len(ps) != 2 {
		return losslessxml.Element{}, fmt.Errorf("paragraph topology")
	}
	if strings.HasPrefix(kind, "paragraph-") {
		props := pptxChild(d, ps[0], pptxElem(pptxA, "pPr"))
		if len(props) != 1 {
			return losslessxml.Element{}, fmt.Errorf("paragraph properties")
		}
		return props[0], nil
	}
	runs := pptxChild(d, ps[0], pptxElem(pptxA, "r"))
	if len(runs) != 2 {
		return losslessxml.Element{}, fmt.Errorf("run topology")
	}
	idx := 0
	if kind == "run-bold" {
		idx = 1
	}
	props := pptxChild(d, runs[idx], pptxElem(pptxA, "rPr"))
	if kind == "run-bold" && len(props) == 0 {
		return runs[idx], nil
	}
	if len(props) != 1 {
		return losslessxml.Element{}, fmt.Errorf("run properties")
	}
	return props[0], nil
}

// Compare the whole source XML except the uniquely selected edit span. A
// missing original rPr permits insertion only immediately before its a:t.
func formattingMaskedSlideSame(before, after []byte, kind string) error {
	old, e := losslessxml.Parse(before)
	if e != nil {
		return e
	}
	saved, e := losslessxml.Parse(after)
	if e != nil {
		return e
	}
	oldEl, e := formattingSelected(old, kind)
	if e != nil {
		return e
	}
	newEl, e := formattingSelected(saved, kind)
	if e != nil {
		return e
	}
	var a, b, c, d int
	if kind == "hide" || kind == "unhide" || strings.HasPrefix(kind, "visibility-") {
		a, b, e = formattingOpeningRange(before, oldEl)
		if e != nil {
			return e
		}
		c, d, e = formattingOpeningRange(after, newEl)
		if e != nil {
			return e
		}
	} else if kind == "run-bold" && oldEl.Name() == pptxElem(pptxA, "r") {
		text := pptxChild(old, oldEl, pptxElem(pptxA, "t"))
		if len(text) != 1 {
			return fmt.Errorf("original run text topology")
		}
		a, _ = text[0].SourceRange()
		b = a
		newShape, err := pptxShape(saved, 4)
		if err != nil {
			return err
		}
		newBody := pptxChild(saved, newShape, pptxElem(pptxP, "txBody"))
		if len(newBody) != 1 {
			return fmt.Errorf("new run body")
		}
		newPs := pptxChild(saved, newBody[0], pptxElem(pptxA, "p"))
		if len(newPs) != 2 {
			return fmt.Errorf("new paragraphs")
		}
		newRuns := pptxChild(saved, newPs[0], pptxElem(pptxA, "r"))
		if len(newRuns) != 2 {
			return fmt.Errorf("new runs")
		}
		props := pptxChild(saved, newRuns[1], pptxElem(pptxA, "rPr"))
		if len(props) != 1 {
			return fmt.Errorf("new run property")
		}
		c, d = props[0].SourceRange()
	} else {
		a, b = oldEl.SourceRange()
		c, d = newEl.SourceRange()
	}
	if a < 0 || b < a || b > len(before) || c < 0 || d < c || d > len(after) {
		return fmt.Errorf("invalid selected property range")
	}
	if !bytes.Equal(before[:a], after[:c]) || !bytes.Equal(before[b:], after[d:]) {
		return fmt.Errorf("nonselected formatting XML bytes changed")
	}
	return formattingUnpatchedSpan(before[a:b], after[c:d], kind)
}

// Mask only contracted attribute values or selected child nodes. Every other
// byte of the selected property (including language, whitespace, child order,
// and unknown direct metadata) must remain literally identical.
func formattingUnpatchedSpan(before, after []byte, kind string) error {
	if kind == "run-bold" && len(before) == 0 {
		if !bytes.Equal(after, []byte(`<a:rPr b="1"/>`)) {
			return fmt.Errorf("new run property contains unproved metadata")
		}
		return nil
	}
	allow := map[string]bool{}
	children := map[string]bool{}
	switch kind {
	case "hide", "unhide", "visibility-order":
		allow["show"] = true
	case "move", "resize", "transform":
		return formattingGeometrySpanSame(before, after, kind)
	case "run-bold", "run-unbold":
		allow["b"] = true
	case "run-italic":
		allow["i"] = true
	case "run-underline":
		allow["u"] = true
	case "run-size":
		allow["sz"] = true
	case "run-font":
		children["latin"] = true
	case "run-color":
		children["solidFill"] = true
	case "run-inherit":
		for _, key := range []string{"b", "i", "u", "sz"} {
			allow[key] = true
		}
		children["latin"] = true
		children["solidFill"] = true
	case "paragraph-align":
		allow["algn"] = true
	case "paragraph-indent":
		allow["marL"] = true
		allow["indent"] = true
	case "paragraph-spacing":
		children["spcBef"] = true
		children["spcAft"] = true
	case "paragraph-line":
		children["lnSpc"] = true
	default:
		return fmt.Errorf("unknown formatting span %s", kind)
	}
	a, e := formattingMaskProperty(before, allow, children)
	if e != nil {
		return e
	}
	b, e := formattingMaskProperty(after, allow, children)
	if e != nil {
		return e
	}
	if !bytes.Equal(a, b) {
		return fmt.Errorf("unpatched selected-property bytes changed")
	}
	return nil
}

type formattingRange struct{ start, end int }

func formattingStrip(raw []byte, ranges []formattingRange) ([]byte, error) {
	for i := range ranges {
		if ranges[i].start < 0 || ranges[i].end < ranges[i].start || ranges[i].end > len(raw) {
			return nil, fmt.Errorf("invalid formatting mask range")
		}
		for j := 0; j < i; j++ {
			if ranges[i].start < ranges[j].end && ranges[j].start < ranges[i].end {
				return nil, fmt.Errorf("overlapping formatting mask")
			}
		}
	}
	for i := 0; i < len(ranges); i++ {
		for j := i + 1; j < len(ranges); j++ {
			if ranges[i].start < ranges[j].start {
				ranges[i], ranges[j] = ranges[j], ranges[i]
			}
		}
	}
	out := append([]byte{}, raw...)
	for _, r := range ranges {
		out = append(append([]byte{}, out[:r.start]...), out[r.end:]...)
	}
	return out, nil
}

// Independent quote-consuming scanner: never match attribute-like text within
// another quoted value, nor mistake > inside a value for the tag boundary.
func formattingOpeningMasks(raw []byte, allow map[string]bool) (int, []formattingRange, error) {
	if len(raw) < 3 || raw[0] != '<' {
		return 0, nil, fmt.Errorf("opening missing")
	}
	nameChar := func(c byte) bool {
		return c == '_' || c == ':' || c == '-' || c == '.' || c >= '0' && c <= '9' || c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z'
	}
	i := 1
	for i < len(raw) && nameChar(raw[i]) {
		i++
	}
	if i == 1 {
		return 0, nil, fmt.Errorf("opening name missing")
	}
	ranges := []formattingRange{}
	seen := map[string]bool{}
	for i < len(raw) {
		start := i
		for i < len(raw) && (raw[i] == ' ' || raw[i] == '\n' || raw[i] == '\r' || raw[i] == '\t') {
			i++
		}
		if i >= len(raw) {
			break
		}
		if raw[i] == '>' {
			return i + 1, ranges, nil
		}
		if raw[i] == '/' && i+1 < len(raw) && raw[i+1] == '>' {
			return i + 2, ranges, nil
		}
		if start == i {
			return 0, nil, fmt.Errorf("attribute separator")
		}
		nameStart := i
		for i < len(raw) && nameChar(raw[i]) {
			i++
		}
		if i == nameStart {
			return 0, nil, fmt.Errorf("invalid attribute name")
		}
		key := string(raw[nameStart:i])
		if seen[key] {
			return 0, nil, fmt.Errorf("duplicate attribute")
		}
		seen[key] = true
		for i < len(raw) && (raw[i] == ' ' || raw[i] == '\t') {
			i++
		}
		if i >= len(raw) || raw[i] != '=' {
			return 0, nil, fmt.Errorf("attribute assignment")
		}
		i++
		for i < len(raw) && (raw[i] == ' ' || raw[i] == '\t') {
			i++
		}
		if i >= len(raw) || raw[i] != '\x27' && raw[i] != '"' {
			return 0, nil, fmt.Errorf("attribute quote")
		}
		q := raw[i]
		i++
		for i < len(raw) && raw[i] != q {
			i++
		}
		if i >= len(raw) {
			return 0, nil, fmt.Errorf("unterminated attribute")
		}
		i++
		if allow[key] {
			ranges = append(ranges, formattingRange{start, i})
		}
	}
	return 0, nil, fmt.Errorf("unterminated opening")
}
func formattingMaskProperty(raw []byte, allow, children map[string]bool) ([]byte, error) {
	_, ranges, e := formattingOpeningMasks(raw, allow)
	if e != nil {
		return nil, e
	}
	if len(children) > 0 {
		prefix := []byte(`<root xmlns:a="` + pptxA + `">`)
		wrapped := append(append(append([]byte{}, prefix...), raw...), []byte(`</root>`)...)
		doc, e := losslessxml.Parse(wrapped)
		if e != nil {
			return nil, e
		}
		elements := doc.Elements()
		if len(elements) < 2 {
			return nil, fmt.Errorf("selected property missing")
		}
		owner := elements[1]
		seen := map[string]bool{}
		for _, node := range elements {
			parent, ok := node.Parent()
			if !ok || parent != owner || node.Name().Space != pptxA {
				continue
			}
			name := node.Name().Local
			if !children[name] {
				continue
			}
			if seen[name] {
				return nil, fmt.Errorf("duplicate selected child %s", name)
			}
			seen[name] = true
			a, b := node.SourceRange()
			ranges = append(ranges, formattingRange{a - len(prefix), b - len(prefix)})
		}
	}
	return formattingStrip(raw, ranges)
}
func formattingGeometrySpanSame(before, after []byte, kind string) error {
	attrs := map[string]bool{}
	if kind == "transform" {
		for _, key := range []string{"rot", "flipH", "flipV"} {
			attrs[key] = true
		}
	}
	a, e := formattingMaskProperty(before, attrs, map[string]bool{"off": true, "ext": true})
	if e != nil {
		return e
	}
	b, e := formattingMaskProperty(after, attrs, map[string]bool{"off": true, "ext": true})
	if e != nil {
		return e
	}
	if !bytes.Equal(a, b) {
		return fmt.Errorf("unpatched transform metadata changed")
	}
	prefix := []byte(`<root xmlns:a="` + pptxA + `">`)
	read := func(raw []byte) (map[string][]byte, error) {
		wrapped := append(append(append([]byte{}, prefix...), raw...), []byte(`</root>`)...)
		doc, e := losslessxml.Parse(wrapped)
		if e != nil {
			return nil, e
		}
		out := map[string][]byte{}
		for _, node := range doc.Elements() {
			parent, ok := node.Parent()
			if !ok || parent != doc.Elements()[1] || node.Name().Space != pptxA {
				continue
			}
			name := node.Name().Local
			if name != "off" && name != "ext" {
				continue
			}
			if _, exists := out[name]; exists {
				return nil, fmt.Errorf("duplicate geometry child")
			}
			start, end := node.SourceRange()
			out[name] = append([]byte{}, raw[start-len(prefix):end-len(prefix)]...)
		}
		if len(out) != 2 {
			return nil, fmt.Errorf("geometry children absent")
		}
		return out, nil
	}
	x, e := read(before)
	if e != nil {
		return e
	}
	y, e := read(after)
	if e != nil {
		return e
	}
	for _, key := range []string{"off", "ext"} {
		allow := map[string]bool{}
		if kind == "move" && key == "off" {
			allow = map[string]bool{"x": true, "y": true}
		}
		if kind == "resize" && key == "ext" {
			allow = map[string]bool{"cx": true, "cy": true}
		}
		left, e := formattingMaskProperty(x[key], allow, nil)
		if e != nil {
			return e
		}
		right, e := formattingMaskProperty(y[key], allow, nil)
		if e != nil {
			return e
		}
		if !bytes.Equal(left, right) {
			return fmt.Errorf("unpatched geometry %s changed", key)
		}
	}
	return nil
}
func formattingSlideListSame(before, after []byte) error {
	old, e := losslessxml.Parse(before)
	if e != nil {
		return e
	}
	saved, e := losslessxml.Parse(after)
	if e != nil {
		return e
	}
	list := func(d *losslessxml.Document) (losslessxml.Element, error) {
		var found losslessxml.Element
		count := 0
		for _, n := range d.Elements() {
			if n.Name() == pptxElem(pptxP, "sldIdLst") {
				found = n
				count++
			}
		}
		if count != 1 {
			return found, fmt.Errorf("slide list count")
		}
		return found, nil
	}
	o, e := list(old)
	if e != nil {
		return e
	}
	n, e := list(saved)
	if e != nil {
		return e
	}
	a, b := o.ContentRange()
	c, d := n.ContentRange()
	if !bytes.Equal(before[:a], after[:c]) || !bytes.Equal(before[b:], after[d:]) {
		return fmt.Errorf("presentation XML outside slide list changed")
	}
	ids := func(doc *losslessxml.Document, el losslessxml.Element, source []byte) [][]byte {
		out := [][]byte{}
		for _, item := range pptxChild(doc, el, pptxElem(pptxP, "sldId")) {
			x, y := item.SourceRange()
			out = append(out, source[x:y])
		}
		return out
	}
	oldEntries, newEntries := ids(old, o, before), ids(saved, n, after)
	if len(oldEntries) != 3 || len(newEntries) != 3 {
		return fmt.Errorf("slide list entry count")
	}
	for i, j := range []int{2, 0, 1} {
		if !bytes.Equal(newEntries[i], oldEntries[j]) {
			return fmt.Errorf("slide entry lexical drift %d", i)
		}
	}
	return nil
}
