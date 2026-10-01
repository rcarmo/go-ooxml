package acceptance

import (
	"bytes"
	"encoding/json"
	"fmt"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const retainedDocxW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

func retainedChild(d *losslessxml.Document, p losslessxml.Element, ns, local string) []losslessxml.Element {
	out := []losslessxml.Element{}
	for _, n := range d.Elements() {
		parent, ok := n.Parent()
		if ok && parent == p && n.Name().Space == ns && n.Name().Local == local {
			out = append(out, n)
		}
	}
	return out
}
func retainedSelect(d *losslessxml.Document, r retainedRecord) (losslessxml.Element, losslessxml.Element, error) {
	if len(d.Elements()) == 0 {
		return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("empty selected XML")
	}
	if r.Format == "pptx" {
		shape, e := pptxShape(d, 4)
		if e != nil {
			return losslessxml.Element{}, losslessxml.Element{}, e
		}
		sp := retainedChild(d, shape, pptxP, "spPr")
		if len(sp) != 1 {
			return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("spPr count")
		}
		switch {
		case strings.HasPrefix(r.Kind, "shape-") || r.Kind == "style-refusal":
			return sp[0], sp[0], nil
		case strings.HasPrefix(r.Kind, "line-"):
			ln := retainedChild(d, sp[0], pptxA, "ln")
			if len(ln) != 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("line count")
			}
			return sp[0], ln[0], nil
		case strings.HasPrefix(r.Kind, "body-"):
			body := retainedChild(d, shape, pptxP, "txBody")
			if len(body) != 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("txBody count")
			}
			prop := retainedChild(d, body[0], pptxA, "bodyPr")
			if len(prop) != 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("bodyPr count")
			}
			return prop[0], prop[0], nil
		case strings.HasPrefix(r.Kind, "run-"):
			body := retainedChild(d, shape, pptxP, "txBody")
			if len(body) != 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("txBody count")
			}
			ps := retainedChild(d, body[0], pptxA, "p")
			if len(ps) < 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("paragraph count")
			}
			rs := retainedChild(d, ps[0], pptxA, "r")
			if len(rs) < 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("run count")
			}
			prop := retainedChild(d, rs[0], pptxA, "rPr")
			if len(prop) != 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("rPr count")
			}
			return prop[0], prop[0], nil
		}
	} else {
		root := d.Elements()[0]
		body := retainedChild(d, root, retainedDocxW, "body")
		if len(body) != 1 {
			return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("Word body count")
		}
		ps := retainedChild(d, body[0], retainedDocxW, "p")
		if len(ps) < 1 {
			return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("Word paragraph count")
		}
		if strings.HasPrefix(r.Kind, "paragraph-") {
			prop := retainedChild(d, ps[0], retainedDocxW, "pPr")
			if len(prop) != 1 {
				return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("Word pPr count")
			}
			return prop[0], prop[0], nil
		}
		rs := retainedChild(d, ps[0], retainedDocxW, "r")
		if len(rs) < 1 {
			return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("Word run count")
		}
		prop := retainedChild(d, rs[0], retainedDocxW, "rPr")
		if len(prop) != 1 {
			return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("Word rPr count")
		}
		return prop[0], prop[0], nil
	}
	return losslessxml.Element{}, losslessxml.Element{}, fmt.Errorf("unknown retained selection %s", r.Kind)
}
func retainedAttr(el losslessxml.Element, ns, local string) string {
	for _, a := range el.Attributes() {
		if a.Name.Space == ns && a.Name.Local == local {
			return a.Value
		}
	}
	return ""
}
func retainedExpected(s *retainedState) error {
	if s.after == nil {
		return fmt.Errorf("no saved archive")
	}
	var want map[string]any
	if e := json.Unmarshal(s.record.Expected, &want); e != nil {
		return e
	}
	if reason, ok := want["refusal"]; ok {
		if s.refusal == nil {
			return fmt.Errorf("missing refusal %v", reason)
		}
		return nil
	}
	doc, e := losslessxml.Parse(s.after[s.part()])
	if e != nil {
		return e
	}
	_, node, e := retainedSelect(doc, s.record)
	if e != nil {
		return e
	}
	ns := pptxA
	if s.record.Format == "docx" {
		ns = retainedDocxW
	}
	value := func(key string) (string, error) {
		v, ok := want[key]
		if !ok {
			return "", fmt.Errorf("missing expected %s", key)
		}
		switch x := v.(type) {
		case string:
			return x, nil
		case bool:
			if x {
				return "1", nil
			}
			return "0", nil
		case float64:
			return strconv.FormatInt(int64(x), 10), nil
		}
		return "", fmt.Errorf("invalid expected %s", key)
	}
	leaf := func(owner losslessxml.Element, name string) (losslessxml.Element, error) {
		xs := retainedChild(doc, owner, ns, name)
		if len(xs) != 1 {
			return losslessxml.Element{}, fmt.Errorf("direct %s count %d", name, len(xs))
		}
		return xs[0], nil
	}
	switch s.record.Kind {
	case "shape-solid-fill", "shape-no-fill", "shape-inherit-fill", "line-color", "line-no-fill":
		if absent, ok := want["absent"]; ok {
			for _, x := range absent.([]any) {
				if len(retainedChild(doc, node, ns, x.(string))) != 0 {
					return fmt.Errorf("direct %s remained", x)
				}
			}
			return nil
		}
		key := "fill"
		if strings.HasPrefix(s.record.Kind, "line-") {
			key = "lineColor"
		}
		v, _ := value(key)
		if v == "none" {
			if len(retainedChild(doc, node, ns, "noFill")) != 1 || len(retainedChild(doc, node, ns, "solidFill")) != 0 {
				return fmt.Errorf("direct noFill choice")
			}
			return nil
		}
		fill, er := leaf(node, "solidFill")
		if er != nil {
			return er
		}
		rgb, er := leaf(fill, "srgbClr")
		if er != nil {
			return er
		}
		if got := retainedAttr(rgb, "", "val"); got != v {
			return fmt.Errorf("sRGB=%q want %q", got, v)
		}
		return nil
	case "line-dash":
		dash, er := leaf(node, "prstDash")
		if er != nil {
			return er
		}
		v, _ := value("lineDash")
		if got := retainedAttr(dash, "", "val"); got != v {
			return fmt.Errorf("dash=%q want %q", got, v)
		}
		return nil
	case "line-width":
		v, _ := value("lineWidth")
		if got := retainedAttr(node, "", "w"); got != v {
			return fmt.Errorf("line width=%q want %q", got, v)
		}
		return nil
	case "paragraph-spacing", "paragraph-line", "paragraph-indent":
		child := "spacing"
		if s.record.Kind == "paragraph-indent" {
			child = "ind"
		}
		n, er := leaf(node, child)
		if er != nil {
			return er
		}
		for key := range want {
			v, er := value(key)
			if er != nil {
				return er
			}
			if got := retainedAttr(n, ns, key); got != v {
				return fmt.Errorf("Word %s=%q want %q", key, got, v)
			}
		}
		return nil
	case "run-font":
		n, er := leaf(node, "rFonts")
		if er != nil {
			return er
		}
		for key := range want {
			v, _ := value(key)
			if got := retainedAttr(n, ns, key); got != v {
				return fmt.Errorf("font %s=%q want %q", key, got, v)
			}
		}
		return nil
	case "run-inherit":
		for _, raw := range want["absent"].([]any) {
			if len(retainedChild(doc, node, ns, raw.(string))) != 0 {
				return fmt.Errorf("inherited %s still direct", raw)
			}
		}
		for _, raw := range want["retained"].([]any) {
			if len(retainedChild(doc, node, ns, raw.(string))) != 1 {
				return fmt.Errorf("retained %s lost", raw)
			}
		}
		return nil
	}
	for key := range want {
		v, er := value(key)
		if er != nil {
			return er
		}
		if s.record.Format == "pptx" {
			if got := retainedAttr(node, "", key); got != v {
				return fmt.Errorf("PPTX %s=%q want %q", key, got, v)
			}
			continue
		}
		leafName := key
		switch key {
		case "strike", "underline", "color", "highlight", "verticalAlign", "caps", "smallCaps", "vanish", "sizeHalfPoints", "jc", "keepLines", "keepNext", "pageBreakBefore", "outlineLvl":
			leafName = map[string]string{"underline": "u", "verticalAlign": "vertAlign", "sizeHalfPoints": "sz"}[key]
			if leafName == "" {
				leafName = key
			}
		}
		n, er := leaf(node, leafName)
		if er != nil {
			return er
		}
		if got := retainedAttr(n, ns, "val"); got != v {
			return fmt.Errorf("Word %s=%q want %q", key, got, v)
		}
	}
	return nil
}
func retainedChildOrder(s *retainedState) error {
	doc, e := losslessxml.Parse(s.after[s.part()])
	if e != nil {
		return e
	}
	_, selected, e := retainedSelect(doc, s.record)
	if e != nil {
		return e
	}
	ns := pptxA
	order := map[string]int{}
	if s.record.Format == "docx" {
		ns = retainedDocxW
		if strings.HasPrefix(s.record.Kind, "run-") {
			order = map[string]int{"rFonts": 1, "b": 2, "i": 3, "caps": 4, "smallCaps": 5, "strike": 6, "dstrike": 7, "vanish": 8, "color": 9, "sz": 10, "highlight": 11, "u": 12, "vertAlign": 13, "lang": 14}
		} else {
			order = map[string]int{"keepNext": 1, "keepLines": 2, "pageBreakBefore": 3, "widowControl": 4, "spacing": 5, "ind": 6, "jc": 7, "outlineLvl": 8}
		}
	} else {
		switch {
		case strings.HasPrefix(s.record.Kind, "shape-") || s.record.Kind == "style-refusal":
			order = map[string]int{"xfrm": 1, "prstGeom": 2, "solidFill": 3, "noFill": 3, "ln": 4}
		case strings.HasPrefix(s.record.Kind, "line-"):
			order = map[string]int{"solidFill": 1, "noFill": 1, "prstDash": 2}
		case strings.HasPrefix(s.record.Kind, "run-"):
			order = map[string]int{"solidFill": 1, "latin": 2}
		}
	}
	if s.record.Format == "pptx" && (strings.HasPrefix(s.record.Kind, "shape-") || s.record.Kind == "style-refusal") {
		lines := retainedChild(doc, selected, pptxA, "ln")
		if len(lines) != 1 {
			return fmt.Errorf("selected line count")
		}
		if e := retainedOrderedChildren(doc, lines[0], pptxA, map[string]int{"solidFill": 1, "noFill": 1, "prstDash": 2}); e != nil {
			return e
		}
	}
	if s.record.Format == "pptx" && strings.HasPrefix(s.record.Kind, "body-") {
		body := retainedChild(doc, selected, pptxA, "noAutofit")
		if len(body) != 1 {
			return fmt.Errorf("noAutofit child drift")
		}
	}
	return retainedOrderedChildren(doc, selected, ns, order)
}
func retainedOrderedChildren(doc *losslessxml.Document, selected losslessxml.Element, ns string, order map[string]int) error {
	prev := -1
	seen := map[string]bool{}
	for _, n := range doc.Elements() {
		p, ok := n.Parent()
		if !ok || p != selected || n.Name().Space != ns {
			continue
		}
		rank, exists := order[n.Name().Local]
		if len(order) > 0 && !exists {
			return fmt.Errorf("unknown property child %s", n.Name().Local)
		}
		if rank < prev {
			return fmt.Errorf("property order %s after %d", n.Name().Local, prev)
		}
		if seen[n.Name().Local] {
			return fmt.Errorf("duplicate direct %s", n.Name().Local)
		}
		seen[n.Name().Local] = true
		prev = rank
	}
	return nil
}
func retainedMaskedPartSame(before, after []byte, r retainedRecord) error {
	old, e := losslessxml.Parse(before)
	if e != nil {
		return e
	}
	saved, e := losslessxml.Parse(after)
	if e != nil {
		return e
	}
	oldOuter, oldInner, e := retainedSelect(old, r)
	if e != nil {
		return e
	}
	newOuter, newInner, e := retainedSelect(saved, r)
	if e != nil {
		return e
	}
	a, b := oldOuter.SourceRange()
	c, d := newOuter.SourceRange()
	if !bytes.Equal(before[:a], after[:c]) || !bytes.Equal(before[b:], after[d:]) {
		return fmt.Errorf("nonselected XML drift")
	}
	rawBefore, rawAfter := bytes.Clone(before[a:b]), bytes.Clone(after[c:d])
	var mask struct {
		Owner      string   `json:"owner"`
		Attributes []string `json:"attributes"`
		Children   []string `json:"children"`
	}
	if e = json.Unmarshal(r.Mask, &mask); e != nil {
		return e
	}
	if mask.Owner != "" { // Retain all bytes outside a uniquely selected direct child (line, spacing or indentation).
		ns := pptxA
		if r.Format == "docx" {
			ns = retainedDocxW
		}
		oldNodes := retainedChild(old, oldOuter, ns, mask.Owner)
		newNodes := retainedChild(saved, newOuter, ns, mask.Owner)
		if len(oldNodes) != 1 || len(newNodes) != 1 {
			return fmt.Errorf("masked owner %s count", mask.Owner)
		}
		oldInner, newInner = oldNodes[0], newNodes[0]
		x, y := oldInner.SourceRange()
		u, v := newInner.SourceRange()
		x -= a
		y -= a
		u -= c
		v -= c
		if !bytes.Equal(rawBefore[:x], rawAfter[:u]) || !bytes.Equal(rawBefore[y:], rawAfter[v:]) {
			return fmt.Errorf("nonselected property siblings drift")
		}
		rawBefore = rawBefore[x:y]
		rawAfter = rawAfter[u:v]
	}
	attrs := map[string]bool{}
	children := map[string]bool{}
	for _, name := range mask.Attributes {
		attrs[name] = true
		if r.Format == "docx" {
			attrs["w:"+name] = true
		}
	}
	for _, name := range mask.Children {
		children[name] = true
	}
	if r.Format == "pptx" {
		left, er := retainedMaskPPTX(rawBefore, attrs, children)
		if er != nil {
			return er
		}
		right, er := retainedMaskPPTX(rawAfter, attrs, children)
		if er != nil {
			return er
		}
		if !bytes.Equal(left, right) {
			return fmt.Errorf("unpatched PPTX property bytes drift")
		}
		return nil
	}
	left, er := retainedMaskWord(rawBefore, attrs, children)
	if er != nil {
		return er
	}
	right, er := retainedMaskWord(rawAfter, attrs, children)
	if er != nil {
		return er
	}
	if !bytes.Equal(left, right) {
		return fmt.Errorf("unpatched Word property bytes drift")
	}
	return nil
}
func retainedMaskPPTX(raw []byte, attrs, children map[string]bool) ([]byte, error) {
	_, ranges, e := formattingOpeningMasks(raw, attrs)
	if e != nil {
		return nil, e
	}
	if len(children) > 0 {
		prefix := []byte(`<root xmlns:a="` + pptxA + `" xmlns:p="` + pptxP + `">`)
		doc, er := losslessxml.Parse(append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...))
		if er != nil {
			return nil, er
		}
		owner := doc.Elements()[1]
		seen := map[string]bool{}
		for _, n := range doc.Elements() {
			p, ok := n.Parent()
			if !ok || p != owner || n.Name().Space != pptxA || !children[n.Name().Local] {
				continue
			}
			if seen[n.Name().Local] {
				return nil, fmt.Errorf("duplicate masked PPTX child")
			}
			seen[n.Name().Local] = true
			a, b := n.SourceRange()
			ranges = append(ranges, formattingRange{start: a - len(prefix), end: b - len(prefix)})
		}
	}
	return formattingStrip(raw, ranges)
}
func retainedMaskWord(raw []byte, attrs, children map[string]bool) ([]byte, error) {
	_, ranges, e := formattingOpeningMasks(raw, attrs)
	if e != nil {
		return nil, e
	}
	if len(children) > 0 {
		prefix := []byte(`<root xmlns:w="` + retainedDocxW + `">`)
		doc, er := losslessxml.Parse(append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...))
		if er != nil {
			return nil, er
		}
		owner := doc.Elements()[1]
		seen := map[string]bool{}
		for _, child := range doc.Elements() {
			p, ok := child.Parent()
			if !ok || p != owner || child.Name().Space != retainedDocxW || !children[child.Name().Local] {
				continue
			}
			if seen[child.Name().Local] {
				return nil, fmt.Errorf("duplicate masked Word child")
			}
			seen[child.Name().Local] = true
			a, b := child.SourceRange()
			ranges = append(ranges, formattingRange{start: a - len(prefix), end: b - len(prefix)})
		}
	}
	return formattingStrip(raw, ranges)
}
