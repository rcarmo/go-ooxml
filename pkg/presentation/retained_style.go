package presentation

import (
	"bytes"
	"regexp"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// RetainedShapeStylePatch edits only direct fill and line styling of an owned shape.
// InheritFill removes the direct fill; it does not resolve inherited appearance.
type RetainedShapeStylePatch struct {
	Fill        *string
	InheritFill bool
	LineColor   *string
	LineWidth   *int64
	LineDash    *string
}
type RetainedBodyPatch struct {
	Anchor, Vert, Wrap                          *string
	AnchorCtr, RTLCol                           *bool
	Rot, LIns, TIns, RIns, BIns, NumCol, SpcCol *int64
}
type RetainedRunPatch struct {
	Strike, Cap       *string
	Baseline, Spacing *int64
}

func retainedColor(v string) bool { return regexp.MustCompile(`^[0-9A-F]{6}$`).MatchString(v) }
func retainedChoice(v string) ([]byte, error) {
	if v == "none" {
		return []byte(`<a:noFill/>`), nil
	}
	if !retainedColor(v) {
		return nil, editRefusal("invalid-style", "sRGB colour must have six uppercase hex digits")
	}
	return []byte(`<a:solidFill><a:srgbClr val="` + v + `"/></a:solidFill>`), nil
}
func retainedDoc(raw []byte) (*losslessxml.Document, losslessxml.Element, int, error) {
	prefix := []byte(`<root xmlns:a="` + packaging.NSDrawingML + `" xmlns:p="` + packaging.NSPresentationML + `">`)
	wrapped := append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...)
	d, e := losslessxml.Parse(wrapped)
	if e != nil || len(d.Elements()) < 2 {
		return nil, losslessxml.Element{}, 0, editRefusal("unsupported_structure", "invalid direct property XML")
	}
	return d, d.Elements()[1], len(prefix), nil
}
func retainedDirect(d *losslessxml.Document, parent losslessxml.Element, allowed map[string]bool) (map[string]losslessxml.Element, error) {
	result := map[string]losslessxml.Element{}
	for _, child := range d.Elements() {
		p, ok := child.Parent()
		if !ok || p != parent {
			continue
		}
		n := child.Name()
		if n.Space != packaging.NSDrawingML || !allowed[n.Local] {
			return nil, editRefusal("unsupported_structure", "unknown direct property child")
		}
		if _, exists := result[n.Local]; exists {
			return nil, editRefusal("unsupported_structure", "duplicate direct child")
		}
		result[n.Local] = child
	}
	return result, nil
}
func retainedPatchChild(raw []byte, names []string, replacement []byte, beforeNames []string, allowed map[string]bool) ([]byte, error) {
	d, parent, prefix, e := retainedDoc(raw)
	if e != nil {
		return nil, e
	}
	children, e := retainedDirect(d, parent, allowed)
	if e != nil {
		return nil, e
	}
	count := 0
	var selected losslessxml.Element
	for _, name := range names {
		if el, ok := children[name]; ok {
			count++
			selected = el
		}
	}
	if count > 1 {
		return nil, editRefusal("unsupported_structure", "multiple direct choices")
	}
	if count == 1 {
		a, b := selected.SourceRange()
		a -= prefix
		b -= prefix
		return append(append(bytes.Clone(raw[:a]), replacement...), raw[b:]...), nil
	}
	if len(replacement) == 0 {
		return bytes.Clone(raw), nil
	}
	at := -1
	for _, name := range beforeNames {
		if el, ok := children[name]; ok {
			a, _ := el.SourceRange()
			if at < 0 || a-prefix < at {
				at = a - prefix
			}
		}
	}
	if at < 0 {
		at = bytes.LastIndex(raw, []byte("</"))
		if at < 0 {
			if bytes.HasSuffix(raw, []byte("/>")) {
				opening := fmtOpeningEnd(raw)
				if opening < 3 {
					return nil, editRefusal("unsupported_structure", "invalid self-closing property")
				}
				nameEnd := 1
				for nameEnd < len(raw) && raw[nameEnd] != ' ' && raw[nameEnd] != '\t' && raw[nameEnd] != '\r' && raw[nameEnd] != '\n' && raw[nameEnd] != '/' && raw[nameEnd] != '>' {
					nameEnd++
				}
				if nameEnd == 1 {
					return nil, editRefusal("unsupported_structure", "property QName")
				}
				return append(append(append(bytes.Clone(raw[:opening-2]), '>'), replacement...), []byte("</"+string(raw[1:nameEnd])+">")...), nil
			}
			return nil, editRefusal("unsupported_structure", "missing direct property closing tag")
		}
	}
	return append(append(bytes.Clone(raw[:at]), replacement...), raw[at:]...), nil
}
func retainedValidateFill(d *losslessxml.Document, fill losslessxml.Element) error {
	if len(fill.Attributes()) != 0 {
		return editRefusal("unsupported_structure", "fill attributes")
	}
	if fill.Name().Local == "noFill" {
		for _, n := range d.Elements() {
			if p, ok := n.Parent(); ok && p == fill {
				return editRefusal("unsupported_structure", "noFill children")
			}
		}
		return nil
	}
	rgb := fmtChildren(d, fill, name(packaging.NSDrawingML, "srgbClr"))
	if len(rgb) != 1 {
		return editRefusal("unsupported_structure", "complex colour")
	}
	attrs := rgb[0].Attributes()
	if len(attrs) != 1 || attrs[0].Name.Space != "" || attrs[0].Name.Local != "val" || !retainedColor(attrs[0].Value) {
		return editRefusal("unsupported_structure", "unsupported sRGB attributes")
	}
	for _, n := range d.Elements() {
		if p, ok := n.Parent(); ok && p == fill && n != rgb[0] {
			return editRefusal("unsupported_structure", "complex colour child")
		}
		if p, ok := n.Parent(); ok && p == rgb[0] {
			return editRefusal("unsupported_structure", "sRGB transformation")
		}
	}
	return nil
}
func retainedValidateStyle(raw []byte, parentName string, allowed map[string]bool) error {
	d, p, _, e := retainedDoc(raw)
	if e != nil {
		return e
	}
	if p.Name().Local != parentName {
		return editRefusal("unsupported_structure", "direct style parent")
	}
	attrsAllowed := map[string]bool{}
	if parentName == "ln" {
		attrsAllowed["w"] = true
	}
	if e = retainedAllowedAttrs(p, attrsAllowed); e != nil {
		return e
	}
	children, e := retainedDirect(d, p, allowed)
	if e != nil {
		return e
	}
	if len(children) > len(allowed) {
		return editRefusal("unsupported_structure", "unexpected style children")
	}
	previous := -1
	orders := map[string]int{"xfrm": 1, "prstGeom": 2, "solidFill": 3, "noFill": 3, "ln": 4, "prstDash": 5}
	for _, child := range d.Elements() {
		owner, ok := child.Parent()
		if !ok || owner != p {
			continue
		}
		rank := orders[child.Name().Local]
		if rank < previous {
			return editRefusal("unsupported_structure", "direct style child order")
		}
		previous = rank
	}
	for _, n := range []string{"solidFill", "noFill"} {
		if fill, ok := children[n]; ok {
			if e = retainedValidateFill(d, fill); e != nil {
				return e
			}
		}
	}
	if _, a := children["solidFill"]; a {
		if _, b := children["noFill"]; b {
			return editRefusal("unsupported_structure", "ambiguous fill")
		}
	}
	return nil
}
func (s *EditSession) SetRetainedShapeStyle(t *FormatTarget, p RetainedShapeStylePatch) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	if p.Fill == nil && !p.InheritFill && p.LineColor == nil && p.LineWidth == nil && p.LineDash == nil {
		return editRefusal("invalid-style", "empty patch")
	}
	if p.Fill != nil && p.InheritFill {
		return editRefusal("invalid-style", "conflicting fill")
	}
	// All operands are checked before changing any part, including a late width.
	if p.LineWidth != nil && (*p.LineWidth < 0 || *p.LineWidth > 20116800) {
		return editRefusal("invalid-style", "line width")
	}
	if p.LineDash != nil && *p.LineDash != "dash" {
		return editRefusal("invalid-style", "line dash")
	}
	if p.Fill != nil {
		if _, e := retainedChoice(*p.Fill); e != nil {
			return e
		}
	}
	if p.LineColor != nil {
		if _, e := retainedChoice(*p.LineColor); e != nil {
			return e
		}
	}
	source, e := fmtSource(t)
	if e != nil {
		return e
	}
	sp, e := fmtOne(t.doc, t.shape, name(packaging.NSPresentationML, "spPr"))
	if e != nil {
		return e
	}
	raw, e := fmtRaw(source, sp)
	if e != nil {
		return e
	}
	if e = retainedValidateStyle(raw, "spPr", map[string]bool{"xfrm": true, "prstGeom": true, "solidFill": true, "noFill": true, "ln": true}); e != nil {
		return e
	}
	updated := bytes.Clone(raw)
	if p.Fill != nil || p.InheritFill {
		v := []byte(nil)
		if p.Fill != nil {
			v, _ = retainedChoice(*p.Fill)
		}
		updated, e = retainedPatchChild(updated, []string{"solidFill", "noFill"}, v, []string{"ln"}, map[string]bool{"xfrm": true, "prstGeom": true, "solidFill": true, "noFill": true, "ln": true})
		if e != nil {
			return e
		}
	}
	if p.LineColor != nil || p.LineWidth != nil || p.LineDash != nil {
		d, parent, offset, er := retainedDoc(updated)
		if er != nil {
			return er
		}
		line := fmtChildren(d, parent, name(packaging.NSDrawingML, "ln"))
		if len(line) != 1 {
			return editRefusal("unsupported_structure", "one direct line required")
		}
		a, b := line[0].SourceRange()
		a -= offset
		b -= offset
		lineRaw := bytes.Clone(updated[a:b])
		allowed := map[string]bool{"solidFill": true, "noFill": true, "prstDash": true}
		if er = retainedValidateStyle(lineRaw, "ln", allowed); er != nil {
			return er
		}
		if p.LineWidth != nil {
			width := strconv.FormatInt(*p.LineWidth, 10)
			lineRaw, er = fmtOpening(lineRaw, map[string]*string{"w": &width})
			if er != nil {
				return er
			}
		}
		if p.LineColor != nil {
			v, _ := retainedChoice(*p.LineColor)
			lineRaw, er = retainedPatchChild(lineRaw, []string{"solidFill", "noFill"}, v, []string{"prstDash"}, allowed)
			if er != nil {
				return er
			}
		}
		if p.LineDash != nil {
			lineRaw, er = retainedPatchChild(lineRaw, []string{"prstDash"}, []byte(`<a:prstDash val="dash"/>`), nil, allowed)
			if er != nil {
				return er
			}
		}
		updated = append(append(bytes.Clone(updated[:a]), lineRaw...), updated[b:]...)
	}
	out, e := fmtRangePatch(t.doc, source, sp, updated)
	if e != nil {
		return e
	}
	return s.formatCommit(t, out)
}
func retainedAllowedAttrs(element losslessxml.Element, allowed map[string]bool) error {
	for _, a := range element.Attributes() {
		if a.Name.Space != "" || !allowed[a.Name.Local] {
			return editRefusal("unsupported_structure", "unsupported or namespaced property attribute")
		}
	}
	return nil
}
func (s *EditSession) SetRetainedBody(t *FormatTarget, p RetainedBodyPatch) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	attrs := map[string]*string{}
	for _, item := range []struct {
		key   string
		v     *string
		valid map[string]bool
	}{{"anchor", p.Anchor, map[string]bool{"ctr": true}}, {"vert", p.Vert, map[string]bool{"vert": true}}, {"wrap", p.Wrap, map[string]bool{"none": true}}} {
		if item.v != nil {
			if !item.valid[*item.v] {
				return editRefusal("invalid-style", "invalid body enum")
			}
			attrs[item.key] = item.v
		}
	}
	for _, item := range []struct {
		key string
		v   *bool
	}{{"anchorCtr", p.AnchorCtr}, {"rtlCol", p.RTLCol}} {
		if item.v != nil {
			v := "0"
			if *item.v {
				v = "1"
			}
			attrs[item.key] = &v
		}
	}
	for _, item := range []struct {
		key      string
		v        *int64
		min, max int64
	}{{"rot", p.Rot, -21600000, 21600000}, {"lIns", p.LIns, 0, 100000000}, {"tIns", p.TIns, 0, 100000000}, {"rIns", p.RIns, 0, 100000000}, {"bIns", p.BIns, 0, 100000000}, {"numCol", p.NumCol, 1, 16}, {"spcCol", p.SpcCol, 0, 100000000}} {
		if item.v != nil {
			if *item.v < item.min || *item.v > item.max {
				return editRefusal("invalid-style", "invalid body bound")
			}
			v := strconv.FormatInt(*item.v, 10)
			attrs[item.key] = &v
		}
	}
	if len(attrs) == 0 {
		return editRefusal("invalid-style", "empty body patch")
	}
	body, e := fmtOne(t.doc, t.shape, name(packaging.NSPresentationML, "txBody"))
	if e != nil {
		return e
	}
	prop, e := fmtOne(t.doc, body, name(packaging.NSDrawingML, "bodyPr"))
	if e != nil {
		return e
	}
	if e = retainedAllowedAttrs(prop, map[string]bool{"wrap": true, "anchor": true, "vert": true, "anchorCtr": true, "rtlCol": true, "rot": true, "lIns": true, "tIns": true, "rIns": true, "bIns": true, "numCol": true, "spcCol": true}); e != nil {
		return e
	}
	children := fmtChildren(t.doc, prop, name(packaging.NSDrawingML, "noAutofit"))
	for _, n := range t.doc.Elements() {
		owner, ok := n.Parent()
		if ok && owner == prop && (n.Name() != name(packaging.NSDrawingML, "noAutofit")) {
			return editRefusal("unsupported_structure", "unknown text frame child")
		}
	}
	if len(children) != 1 {
		return editRefusal("unsupported_structure", "unsupported text frame")
	}
	source, e := fmtSource(t)
	if e != nil {
		return e
	}
	raw, e := fmtRaw(source, prop)
	if e != nil {
		return e
	}
	updated, e := fmtOpening(raw, attrs)
	if e != nil {
		return e
	}
	out, e := fmtRangePatch(t.doc, source, prop, updated)
	if e != nil {
		return e
	}
	return s.formatCommit(t, out)
}
func (s *EditSession) SetRetainedRun(t *FormatTarget, paragraph, run int, p RetainedRunPatch) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	changes := map[string]*string{}
	for _, v := range []struct {
		key   string
		value *string
		valid string
	}{{"strike", p.Strike, "sngStrike"}, {"cap", p.Cap, "small"}} {
		if v.value != nil {
			if *v.value != v.valid {
				return editRefusal("invalid-style", "invalid run enum")
			}
			changes[v.key] = v.value
		}
	}
	for _, v := range []struct {
		key      string
		value    *int64
		min, max int64
	}{{"baseline", p.Baseline, -100000, 100000}, {"spc", p.Spacing, -400000, 400000}} {
		if v.value != nil {
			if *v.value < v.min || *v.value > v.max {
				return editRefusal("invalid-style", "invalid run bound")
			}
			text := strconv.FormatInt(*v.value, 10)
			changes[v.key] = &text
		}
	}
	if len(changes) == 0 {
		return editRefusal("invalid-style", "empty run patch")
	}
	_, r, e := formatRun(t.doc, t.shape, paragraph, run)
	if e != nil {
		return e
	}
	rp, e := fmtOne(t.doc, r, name(packaging.NSDrawingML, "rPr"))
	if e != nil {
		return e
	}
	source, e := fmtSource(t)
	if e != nil {
		return e
	}
	raw, e := fmtRaw(source, rp)
	if e != nil {
		return e
	}
	for _, a := range rp.Attributes() {
		if a.Name.Space != "" {
			return editRefusal("unsupported_structure", "foreign run attribute")
		}
	}
	updated, e := formatRPr(raw, changes, nil, nil, false)
	if e != nil {
		return e
	}
	out, e := fmtRangePatch(t.doc, source, rp, updated)
	if e != nil {
		return e
	}
	return s.formatCommit(t, out)
}
