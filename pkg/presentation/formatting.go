package presentation

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// FormatTarget is a generation-bound, direct ordinary shape. It deliberately
// does not use the plain-text FindShape policy: rich rPr is expected here.
type FormatTarget struct {
	session    *EditSession
	generation uint64
	part, hash string
	id         uint32
	doc        *losslessxml.Document
	shape      losslessxml.Element
}

func (s *EditSession) formatCurrent(t *FormatTarget) error {
	if t == nil || t.session != s || t.generation != s.generation {
		return editRefusal("stale_target", "foreign or stale formatting target")
	}
	_, h, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	if h != t.hash {
		return editRefusal("stale_target", "format part changed")
	}
	return nil
}
func (s *EditSession) formatCommit(t *FormatTarget, data []byte) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	b, _, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	if bytes.Equal(b, data) {
		return nil
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: t.part, ExpectedSHA256: t.hash, Data: data}}); e != nil {
		return e
	}
	s.generation++
	return nil
}
func fmtChildren(d *losslessxml.Document, p losslessxml.Element, n xml.Name) []losslessxml.Element {
	return manipulationChildren(d, p, n)
}
func fmtOne(d *losslessxml.Document, p losslessxml.Element, n xml.Name) (losslessxml.Element, error) {
	xs := fmtChildren(d, p, n)
	if len(xs) != 1 {
		return losslessxml.Element{}, editRefusal("unsupported_structure", "one direct "+n.Local+" required")
	}
	return xs[0], nil
}
func (s *EditSession) FindFormatShape(part string, id uint32) (*FormatTarget, error) {
	if !s.slides[part] {
		return nil, editRefusal("missing_target", "slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return nil, e
	}
	b, h, e := s.pkg.Part(part)
	if e != nil {
		return nil, e
	}
	d, e := losslessxml.Parse(b)
	if e != nil {
		return nil, e
	}
	es := d.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, editRefusal("unsupported_structure", "slide root")
	}
	var shape losslessxml.Element
	seen := map[uint64]bool{}
	count := 0
	for _, node := range es {
		if node.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		n, er := strconv.ParseUint(manipulationAttr(node, "id"), 10, 32)
		if er != nil || n == 0 {
			return nil, editRefusal("unsupported_structure", "invalid shape ID")
		}
		if seen[n] {
			return nil, editRefusal("ambiguous_target", "duplicate shape ID")
		}
		seen[n] = true
		if uint32(n) != id {
			continue
		}
		count++
		nv, ok := node.Parent()
		if !ok || nv.Name() != name(packaging.NSPresentationML, "nvSpPr") {
			return nil, editRefusal("unsupported_structure", "ordinary shape required")
		}
		shape, ok = nv.Parent()
		if !ok || shape.Name() != name(packaging.NSPresentationML, "sp") {
			return nil, editRefusal("unsupported_structure", "ordinary shape required")
		}
	}
	if count != 1 {
		return nil, editRefusal("missing_target", "unique shape ID required")
	}
	tree, ok := shape.Parent()
	if !ok || tree.Name() != name(packaging.NSPresentationML, "spTree") {
		return nil, editRefusal("unsupported_structure", "grouped shape")
	}
	common, ok := tree.Parent()
	if !ok || common.Name() != name(packaging.NSPresentationML, "cSld") {
		return nil, editRefusal("unsupported_structure", "shape tree owner")
	}
	root, ok := common.Parent()
	if !ok || root != es[0] {
		return nil, editRefusal("unsupported_structure", "slide owner")
	}
	for _, node := range es {
		if node != shape && !manipulationWithin(node, shape) {
			continue
		}
		switch node.Name().Local {
		case "spLocks":
			attrs := node.Attributes()
			if len(attrs) != 1 || attrs[0].Name != (xml.Name{Local: "noGrp"}) || attrs[0].Value != "1" {
				return nil, editRefusal("protected_operation", "locked shape")
			}
		case "fld", "hlinkClick", "hlinkMouseOver", "extLst", "AlternateContent":
			return nil, editRefusal("unsupported_structure", "complex shape")
		}
	}
	return &FormatTarget{s, s.generation, part, h, id, d, shape}, nil
}
func fmtSplice(d *losslessxml.Document, original []byte, start, end int, replacement []byte) ([]byte, error) {
	if start < 0 || end < start || end > len(original) {
		return nil, editRefusal("unsupported_structure", "invalid XML range")
	}
	out := make([]byte, 0, len(original)-end+start+len(replacement))
	out = append(out, original[:start]...)
	out = append(out, replacement...)
	out = append(out, original[end:]...)
	if _, e := losslessxml.Parse(out); e != nil {
		return nil, editRefusal("unsupported_structure", e.Error())
	}
	return out, nil
}
func fmtSource(t *FormatTarget) ([]byte, error) { b, _, e := t.session.pkg.Part(t.part); return b, e }
func fmtRaw(source []byte, e losslessxml.Element) ([]byte, error) {
	a, b := e.SourceRange()
	if a < 0 || b < a || b > len(source) {
		return nil, editRefusal("unsupported_structure", "invalid element range")
	}
	return source[a:b], nil
}

func fmtOpeningEnd(raw []byte) int {
	quote := byte(0)
	for i, c := range raw {
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
			return i + 1
		}
	}
	return -1
}

// fmtOpening changes only real attributes of the opening tag. Attribute-like
// text inside quoted values must never become a second edit target.
func fmtOpening(raw []byte, changes map[string]*string) ([]byte, error) {
	end := -1
	quote := byte(0)
	for i, c := range raw {
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
			end = i
			break
		}
	}
	if end < 0 {
		return nil, editRefusal("unsupported_structure", "missing opening tag")
	}
	tag := raw[:end+1]
	seen := map[string]bool{}
	type span struct {
		a, b  int
		value []byte
	}
	edits := []span{}
	nameByte := func(c byte) bool {
		return c == '_' || c == ':' || c == '-' || c == '.' || c >= '0' && c <= '9' || c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z'
	}
	i := 1
	for i < end && nameByte(tag[i]) {
		i++
	}
	if i == 1 {
		return nil, editRefusal("unsupported_structure", "invalid opening tag")
	}
	for i < end {
		start := i
		for i < end && (tag[i] == ' ' || tag[i] == '\n' || tag[i] == '\r' || tag[i] == '\t') {
			i++
		}
		if i == end || tag[i] == '/' && i == end-1 {
			break
		}
		if i == start {
			return nil, editRefusal("unsupported_structure", "attribute separator required")
		}
		nameStart := i
		for i < end && nameByte(tag[i]) {
			i++
		}
		if i == nameStart {
			return nil, editRefusal("unsupported_structure", "invalid attribute")
		}
		key := string(tag[nameStart:i])
		if seen[key] {
			return nil, editRefusal("ambiguous_target", "duplicate XML attribute")
		}
		seen[key] = true
		for i < end && (tag[i] == ' ' || tag[i] == '\t') {
			i++
		}
		if i >= end || tag[i] != '=' {
			return nil, editRefusal("unsupported_structure", "attribute assignment required")
		}
		i++
		for i < end && (tag[i] == ' ' || tag[i] == '\t') {
			i++
		}
		if i >= end || tag[i] != '\'' && tag[i] != '"' {
			return nil, editRefusal("unsupported_structure", "quoted attribute required")
		}
		q := tag[i]
		i++
		valueStart := i
		for i < end && tag[i] != q {
			i++
		}
		if i >= end {
			return nil, editRefusal("unsupported_structure", "unterminated attribute")
		}
		valueEnd := i
		i++
		if v, ok := changes[key]; ok {
			if v == nil {
				edits = append(edits, span{start, i, nil})
			} else {
				value := []byte(xmlEscapeAttr(*v))
				if q == '\'' {
					value = []byte(strings.ReplaceAll(string(value), "'", "&apos;"))
				}
				if !bytes.Equal(tag[valueStart:valueEnd], value) {
					edits = append(edits, span{valueStart, valueEnd, value})
				}
			}
		}
	}
	for k, v := range changes {
		if seen[k] || v == nil {
			continue
		}
		if strings.Contains(k, ":") {
			return nil, editRefusal("unsupported_structure", "prefixed property")
		}
		at := end
		if at > 0 && tag[at-1] == '/' {
			at--
		}
		edits = append(edits, span{at, at, []byte(` ` + k + `="` + xmlEscapeAttr(*v) + `"`)})
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	out := append([]byte{}, raw...)
	for _, x := range edits {
		out = append(append(append([]byte{}, out[:x.a]...), x.value...), out[x.b:]...)
	}
	return out, nil
}
func xmlEscapeAttr(v string) string {
	var b bytes.Buffer
	_ = xml.EscapeText(&b, []byte(v))
	return strings.ReplaceAll(b.String(), `"`, `&quot;`)
}
func fmtString(s string) *string { return &s }
func fmtRangePatch(d *losslessxml.Document, source []byte, e losslessxml.Element, patch []byte) ([]byte, error) {
	a, b := e.SourceRange()
	return fmtSplice(d, source, a, b, patch)
}

// SetSlideVisibility writes the direct, unqualified CT_Slide show property.
// It does not infer rendered or Office-confirmed hidden state.
func (s *EditSession) SetSlideVisibility(part string, visible bool) error {
	b, h, d, root, e := s.visibilitySource(part)
	if e != nil {
		return e
	}
	v := "0"
	if visible {
		v = "1"
	}
	raw, e := fmtRaw(b, root)
	if e != nil {
		return e
	}
	updated, e := fmtOpening(raw, map[string]*string{"show": &v})
	if e != nil {
		return e
	}
	out, e := fmtRangePatch(d, b, root, updated)
	if e != nil {
		return e
	}
	if bytes.Equal(b, out) {
		return nil
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: out}}); e != nil {
		return e
	}
	s.generation++
	return nil
}
func (s *EditSession) SlideVisible(part string) (bool, error) {
	_, _, _, root, e := s.visibilitySource(part)
	if e != nil {
		return false, e
	}
	v := manipulationAttr(root, "show")
	return v == "" || v == "1" || v == "true", nil
}
func (s *EditSession) visibilitySource(part string) ([]byte, string, *losslessxml.Document, losslessxml.Element, error) {
	if !s.slides[part] {
		return nil, "", nil, losslessxml.Element{}, editRefusal("missing_target", "slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return nil, "", nil, losslessxml.Element{}, e
	}
	b, h, e := s.pkg.Part(part)
	if e != nil {
		return nil, "", nil, losslessxml.Element{}, e
	}
	d, e := losslessxml.Parse(b)
	if e != nil {
		return nil, "", nil, losslessxml.Element{}, e
	}
	es := d.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, "", nil, losslessxml.Element{}, editRefusal("unsupported_structure", "slide root")
	}
	root := es[0]
	seen := false
	for _, a := range root.Attributes() {
		if a.Name.Local != "show" {
			continue
		}
		if a.Name.Space != "" || seen {
			return nil, "", nil, root, editRefusal("invalid-visibility", "ambiguous visibility attribute")
		}
		seen = true
		if a.Value != "0" && a.Value != "1" && a.Value != "false" && a.Value != "true" {
			return nil, "", nil, root, editRefusal("invalid-visibility", "malformed visibility")
		}
	}
	return b, h, d, root, nil
}

// SetSlideVisibilityValue is the typed boundary for untrusted callers.
func (s *EditSession) SetSlideVisibilityValue(part string, v any) error {
	b, ok := v.(bool)
	if !ok {
		return editRefusal("invalid-visibility", "visibility requires Boolean")
	}
	return s.SetSlideVisibility(part, b)
}

type GeometryPatch struct {
	X, Y, Width, Height, Rotation *int64
	FlipH, FlipV                  *bool
}

func (s *EditSession) SetFormatGeometry(t *FormatTarget, p GeometryPatch) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	if p.X == nil && p.Y == nil && p.Width == nil && p.Height == nil && p.Rotation == nil && p.FlipH == nil && p.FlipV == nil {
		return editRefusal("invalid-geometry", "empty patch")
	}
	d := t.doc
	sppr, e := fmtOne(d, t.shape, name(packaging.NSPresentationML, "spPr"))
	if e != nil {
		return e
	}
	xfrm, e := fmtOne(d, sppr, name(packaging.NSDrawingML, "xfrm"))
	if e != nil {
		return e
	}
	off, e := fmtOne(d, xfrm, name(packaging.NSDrawingML, "off"))
	if e != nil {
		return e
	}
	ext, e := fmtOne(d, xfrm, name(packaging.NSDrawingML, "ext"))
	if e != nil {
		return e
	}
	if len(fmtChildren(d, xfrm, name(packaging.NSDrawingML, "chOff"))) > 0 || len(fmtChildren(d, xfrm, name(packaging.NSDrawingML, "chExt"))) > 0 {
		return editRefusal("unsupported_structure", "transformed shape")
	}
	for _, node := range d.Elements() {
		if parent, ok := node.Parent(); ok && parent == xfrm && node != off && node != ext {
			return editRefusal("unsupported_structure", "unknown transform child")
		}
	}
	parse := func(el losslessxml.Element, k string) (int64, error) {
		n, err := strconv.ParseInt(manipulationAttr(el, k), 10, 64)
		if err != nil {
			return 0, editRefusal("unsupported_structure", "invalid source geometry")
		}
		return n, nil
	}
	x, e := parse(off, "x")
	if e != nil {
		return e
	}
	y, e := parse(off, "y")
	if e != nil {
		return e
	}
	w, e := parse(ext, "cx")
	if e != nil {
		return e
	}
	h, e := parse(ext, "cy")
	if e != nil {
		return e
	}
	if p.X != nil {
		x = *p.X
	}
	if p.Y != nil {
		y = *p.Y
	}
	if p.Width != nil {
		w = *p.Width
	}
	if p.Height != nil {
		h = *p.Height
	}
	if x < 0 || y < 0 || w < 1 || h < 1 || x > 100000000 || y > 100000000 || w > 100000000 || h > 100000000 || x > 100000000-w || y > 100000000-h {
		return editRefusal("invalid-geometry", "out of bounds")
	}
	if p.Rotation != nil && (*p.Rotation < 0 || *p.Rotation > 21599999) {
		return editRefusal("invalid-geometry", "rotation out of bounds")
	}
	source, e := fmtSource(t)
	if e != nil {
		return e
	}
	patches := []struct {
		el     losslessxml.Element
		values map[string]*string
	}{}
	atOff := map[string]*string{}
	atExt := map[string]*string{}
	if p.X != nil {
		atOff["x"] = fmtString(strconv.FormatInt(x, 10))
	}
	if p.Y != nil {
		atOff["y"] = fmtString(strconv.FormatInt(y, 10))
	}
	if p.Width != nil {
		atExt["cx"] = fmtString(strconv.FormatInt(w, 10))
	}
	if p.Height != nil {
		atExt["cy"] = fmtString(strconv.FormatInt(h, 10))
	}
	if len(atOff) > 0 {
		patches = append(patches, struct {
			el     losslessxml.Element
			values map[string]*string
		}{off, atOff})
	}
	if len(atExt) > 0 {
		patches = append(patches, struct {
			el     losslessxml.Element
			values map[string]*string
		}{ext, atExt})
	}
	atTransform := map[string]*string{}
	if p.Rotation != nil {
		atTransform["rot"] = fmtString(strconv.FormatInt(*p.Rotation, 10))
	}
	if p.FlipH != nil {
		v := "0"
		if *p.FlipH {
			v = "1"
		}
		atTransform["flipH"] = &v
	}
	if p.FlipV != nil {
		v := "0"
		if *p.FlipV {
			v = "1"
		}
		atTransform["flipV"] = &v
	}
	if len(atTransform) > 0 {
		patches = append(patches, struct {
			el     losslessxml.Element
			values map[string]*string
		}{xfrm, atTransform})
	} // opening tags are disjoint, ordered from right to left
	for i := 0; i < len(patches); i++ {
		for j := i + 1; j < len(patches); j++ {
			a, _ := patches[i].el.SourceRange()
			b, _ := patches[j].el.SourceRange()
			if a < b {
				patches[i], patches[j] = patches[j], patches[i]
			}
		}
	}
	out := append([]byte{}, source...)
	for _, item := range patches {
		a, _ := item.el.SourceRange()
		open, e := fmtOpening(source[a:], item.values)
		if e != nil {
			return e
		} // fmtOpening received the tail; retain only its updated opening tag
		oldEnd := fmtOpeningEnd(source[a:])
		newEnd := fmtOpeningEnd(open)
		if oldEnd < 1 || newEnd < 1 {
			return editRefusal("unsupported_structure", "invalid transform opening")
		}
		out = append(append(append([]byte{}, out[:a]...), open[:newEnd]...), out[a+oldEnd:]...)
	}
	if _, e = losslessxml.Parse(out); e != nil {
		return editRefusal("unsupported_structure", e.Error())
	}
	return s.formatCommit(t, out)
}

// RunFormatPatch applies direct properties, not effective inherited appearance.
type RunFormatPatch struct {
	Bold, Italic *bool
	Underline    *string
	Size         *int64
	Latin, Color *string
	RemoveDirect bool
}

func (s *EditSession) SetFormatRun(t *FormatTarget, paragraph, run int, p RunFormatPatch) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	d := t.doc
	para, r, e := formatRun(d, t.shape, paragraph, run)
	if e != nil {
		return e
	}
	_ = para
	source, e := fmtSource(t)
	if e != nil {
		return e
	}
	rs := fmtChildren(d, r, name(packaging.NSDrawingML, "rPr"))
	if len(rs) > 1 {
		return editRefusal("unsupported_structure", "duplicate run properties")
	}
	if p.RemoveDirect && (p.Bold != nil || p.Italic != nil || p.Underline != nil || p.Size != nil || p.Latin != nil || p.Color != nil) {
		return editRefusal("unsupported_structure", "mixed removal patch")
	}
	changes := map[string]*string{}
	for _, kv := range []struct {
		key   string
		value *bool
	}{{"b", p.Bold}, {"i", p.Italic}} {
		if kv.value != nil {
			v := "0"
			if *kv.value {
				v = "1"
			}
			changes[kv.key] = &v
		}
	}
	if p.Underline != nil {
		if *p.Underline != "sng" && *p.Underline != "none" {
			return editRefusal("unsupported_structure", "underline choice")
		}
		changes["u"] = p.Underline
	}
	if p.Size != nil {
		if *p.Size < 100 || *p.Size > 400000 {
			return editRefusal("unsupported_structure", "font size bounds")
		}
		changes["sz"] = fmtString(strconv.FormatInt(*p.Size, 10))
	}
	if p.Latin != nil {
		if !utf8.ValidString(*p.Latin) || utf8.RuneCountInString(*p.Latin) < 1 || utf8.RuneCountInString(*p.Latin) > 128 || strings.ContainsAny(*p.Latin, "\r\n\t") {
			return editRefusal("unsupported_structure", "Latin typeface bounds")
		}
	}
	if p.Color != nil && !regexp.MustCompile(`^[0-9A-F]{6}$`).MatchString(*p.Color) {
		return editRefusal("unsupported_structure", "sRGB color")
	}
	if len(changes) == 0 && p.Latin == nil && p.Color == nil && !p.RemoveDirect {
		return editRefusal("unsupported_structure", "empty run patch")
	}
	if len(rs) == 0 {
		if p.RemoveDirect {
			return nil
		}
		texts := fmtChildren(d, r, name(packaging.NSDrawingML, "t"))
		if len(texts) != 1 {
			return editRefusal("unsupported_structure", "run text topology")
		}
		a, _ := texts[0].SourceRange()
		raw := []byte(`<a:rPr/>`)
		updated, e := formatRPr(raw, changes, p.Latin, p.Color, false)
		if e != nil {
			return e
		}
		out, e := fmtSplice(d, source, a, a, updated)
		if e != nil {
			return e
		}
		return s.formatCommit(t, out)
	}
	rpr := rs[0]
	raw, e := fmtRaw(source, rpr)
	if e != nil {
		return e
	}
	updated, e := formatRPr(raw, changes, p.Latin, p.Color, p.RemoveDirect)
	if e != nil {
		return e
	}
	out, e := fmtRangePatch(d, source, rpr, updated)
	if e != nil {
		return e
	}
	return s.formatCommit(t, out)
}
func formatRun(d *losslessxml.Document, shape losslessxml.Element, paragraph, run int) (losslessxml.Element, losslessxml.Element, error) {
	body, e := fmtOne(d, shape, name(packaging.NSPresentationML, "txBody"))
	if e != nil {
		return losslessxml.Element{}, losslessxml.Element{}, e
	}
	ps := fmtChildren(d, body, name(packaging.NSDrawingML, "p"))
	if paragraph < 0 || paragraph >= len(ps) {
		return losslessxml.Element{}, losslessxml.Element{}, editRefusal("missing_target", "paragraph index")
	}
	runs := fmtChildren(d, ps[paragraph], name(packaging.NSDrawingML, "r"))
	if run < 0 || run >= len(runs) {
		return losslessxml.Element{}, losslessxml.Element{}, editRefusal("missing_target", "run index")
	}
	if len(fmtChildren(d, runs[run], name(packaging.NSDrawingML, "t"))) != 1 {
		return losslessxml.Element{}, losslessxml.Element{}, editRefusal("unsupported_structure", "run text topology")
	}
	for _, n := range d.Elements() {
		if n != runs[run] && !manipulationWithin(n, runs[run]) {
			continue
		}
		switch n.Name().Local {
		case "fld", "br", "extLst", "hlinkClick", "hlinkMouseOver":
			return losslessxml.Element{}, losslessxml.Element{}, editRefusal("unsupported_structure", "complex text run")
		}
	}
	return ps[paragraph], runs[run], nil
}
func formatRPr(raw []byte, changes map[string]*string, latin, color *string, remove bool) ([]byte, error) {
	prefix := []byte(`<root xmlns:a="` + packaging.NSDrawingML + `">`)
	wrapped := make([]byte, 0, len(prefix)+len(raw)+len(`</root>`))
	wrapped = append(wrapped, prefix...)
	wrapped = append(wrapped, raw...)
	wrapped = append(wrapped, []byte(`</root>`)...)
	d, e := losslessxml.Parse(wrapped)
	if e != nil {
		return nil, editRefusal("unsupported_structure", e.Error())
	}
	if len(d.Elements()) < 2 {
		return nil, editRefusal("unsupported_structure", "run property")
	}
	rpr := d.Elements()[1]
	if rpr.Name() != name(packaging.NSDrawingML, "rPr") {
		return nil, editRefusal("unsupported_structure", "run property")
	}
	for _, a := range rpr.Attributes() {
		if a.Name.Space != "" {
			return nil, editRefusal("unsupported_structure", "foreign run attribute")
		}
	}
	if remove {
		changes = map[string]*string{"b": nil, "i": nil, "u": nil, "sz": nil}
	}
	updated, e := fmtOpening(raw, changes)
	if e != nil {
		return nil, e
	} // patch children with exact original ranges, in descending offset order
	type edit struct {
		a, b int
		data []byte
	}
	edits := []edit{}
	fill, fonts := []losslessxml.Element{}, []losslessxml.Element{}
	for _, n := range d.Elements() {
		if parent, ok := n.Parent(); ok && parent == rpr {
			switch n.Name() {
			case name(packaging.NSDrawingML, "solidFill"):
				fill = append(fill, n)
			case name(packaging.NSDrawingML, "latin"):
				fonts = append(fonts, n)
			default:
				return nil, editRefusal("unsupported_structure", "unknown run property child")
			}
		}
	}
	if len(fill) > 1 || len(fonts) > 1 {
		return nil, editRefusal("unsupported_structure", "ambiguous run property")
	}
	if len(fill) == 1 && len(fonts) == 1 {
		a, _ := fill[0].SourceRange()
		b, _ := fonts[0].SourceRange()
		if a > b {
			return nil, editRefusal("unsupported_structure", "run property child order")
		}
	}
	if len(fill) > 0 {
		cs := fmtChildren(d, fill[0], name(packaging.NSDrawingML, "srgbClr"))
		if len(cs) != 1 || len(fmtChildren(d, fill[0], name(packaging.NSDrawingML, "schemeClr"))) > 0 {
			return nil, editRefusal("unsupported_structure", "complex color")
		}
		for _, child := range d.Elements() {
			if parent, ok := child.Parent(); ok && parent == fill[0] && child != cs[0] {
				return nil, editRefusal("unsupported_structure", "complex color child")
			}
			if parent, ok := child.Parent(); ok && parent == cs[0] {
				return nil, editRefusal("unsupported_structure", "sRGB transformation unsupported")
			}
		}
	}
	// Reparse after opening edit to get valid offsets for child edits.
	wrapped = append(append(append([]byte{}, prefix...), updated...), []byte(`</root>`)...)
	d, e = losslessxml.Parse(wrapped)
	if e != nil {
		return nil, editRefusal("unsupported_structure", e.Error())
	}
	rpr = d.Elements()[1]
	prefixLen := len(prefix)
	fill = fmtChildren(d, rpr, name(packaging.NSDrawingML, "solidFill"))
	fonts = fmtChildren(d, rpr, name(packaging.NSDrawingML, "latin"))
	del := func(el losslessxml.Element, repl []byte) {
		a, b := el.SourceRange()
		edits = append(edits, edit{a - prefixLen, b - prefixLen, repl})
	}
	if len(fill) == 1 && (color != nil || remove) {
		val := []byte{}
		if color != nil && !remove {
			val = []byte(`<a:solidFill><a:srgbClr val="` + *color + `"/></a:solidFill>`)
		}
		del(fill[0], val)
	}
	if len(fonts) == 1 && (latin != nil || remove) {
		val := []byte{}
		if latin != nil && !remove {
			val = []byte(`<a:latin typeface="` + xmlEscapeAttr(*latin) + `"/>`)
		}
		del(fonts[0], val)
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	for _, x := range edits {
		updated = append(append(append([]byte{}, updated[:x.a]...), x.data...), updated[x.b:]...)
	}
	if remove {
		return updated, nil
	}
	needColor := color != nil && len(fill) == 0
	needLatin := latin != nil && len(fonts) == 0
	if !needColor && !needLatin {
		return updated, nil
	}
	addition := []byte{}
	if needColor {
		addition = append(addition, []byte(`<a:solidFill><a:srgbClr val="`+*color+`"/></a:solidFill>`)...)
	}
	if needLatin {
		addition = append(addition, []byte(`<a:latin typeface="`+xmlEscapeAttr(*latin)+`"/>`)...)
	}
	if bytes.HasSuffix(updated, []byte("/>")) {
		return append(append([]byte{}, updated[:len(updated)-2]...), append([]byte(">"), append(addition, []byte("</a:rPr>")...)...)...), nil
	}
	end := bytes.Index(updated, []byte("</a:rPr>"))
	if end < 0 {
		return nil, editRefusal("unsupported_structure", "run property closing")
	}
	if needColor && len(fonts) == 1 { // fill must precede existing latin
		idx := bytes.Index(updated, []byte("<a:latin"))
		if idx < 0 {
			return nil, editRefusal("unsupported_structure", "Latin prefix")
		}
		updated = append(append(append([]byte{}, updated[:idx]...), addition...), updated[idx:]...)
		return updated, nil
	}
	return append(append(append([]byte{}, updated[:end]...), addition...), updated[end:]...), nil
}

type ParagraphFormatPatch struct {
	Alignment                                                *string
	MarginLeft, Indent, SpaceBefore, SpaceAfter, LinePercent *int64
}

func (s *EditSession) SetFormatParagraph(t *FormatTarget, index int, p ParagraphFormatPatch) error {
	if e := s.formatCurrent(t); e != nil {
		return e
	}
	d := t.doc
	body, e := fmtOne(d, t.shape, name(packaging.NSPresentationML, "txBody"))
	if e != nil {
		return e
	}
	ps := fmtChildren(d, body, name(packaging.NSDrawingML, "p"))
	if index < 0 || index >= len(ps) {
		return editRefusal("missing_target", "paragraph index")
	}
	para := ps[index]
	if p.Alignment == nil && p.MarginLeft == nil && p.Indent == nil && p.SpaceBefore == nil && p.SpaceAfter == nil && p.LinePercent == nil {
		return editRefusal("unsupported_structure", "empty paragraph patch")
	}
	changes := map[string]*string{}
	if p.Alignment != nil {
		switch *p.Alignment {
		case "l", "ctr", "r", "just":
		default:
			return editRefusal("unsupported_structure", "paragraph alignment")
		}
		changes["algn"] = p.Alignment
	}
	for _, item := range []struct {
		k        string
		v        *int64
		min, max int64
	}{{"marL", p.MarginLeft, 0, 100000000}, {"indent", p.Indent, -100000000, 100000000}} {
		if item.v != nil {
			if *item.v < item.min || *item.v > item.max {
				return editRefusal("unsupported_structure", "paragraph geometry")
			}
			changes[item.k] = fmtString(strconv.FormatInt(*item.v, 10))
		}
	}
	for _, item := range []struct {
		v        *int64
		min, max int64
	}{{p.SpaceBefore, 0, 158400}, {p.SpaceAfter, 0, 158400}, {p.LinePercent, 1, 13200000}} {
		if item.v != nil && (*item.v < item.min || *item.v > item.max) {
			return editRefusal("unsupported_structure", "paragraph spacing")
		}
	}
	source, e := fmtSource(t)
	if e != nil {
		return e
	}
	props := fmtChildren(d, para, name(packaging.NSDrawingML, "pPr"))
	if len(props) > 1 {
		return editRefusal("unsupported_structure", "duplicate paragraph properties")
	}
	if len(props) == 0 {
		children := fmtChildren(d, para, name(packaging.NSDrawingML, "r"))
		if len(children) == 0 {
			return editRefusal("unsupported_structure", "paragraph without runs")
		}
		a, _ := children[0].SourceRange()
		r, e := formatPPr([]byte(`<a:pPr/>`), changes, p)
		if e != nil {
			return e
		}
		out, e := fmtSplice(d, source, a, a, r)
		if e != nil {
			return e
		}
		return s.formatCommit(t, out)
	}
	raw, e := fmtRaw(source, props[0])
	if e != nil {
		return e
	}
	updated, e := formatPPr(raw, changes, p)
	if e != nil {
		return e
	}
	out, e := fmtRangePatch(d, source, props[0], updated)
	if e != nil {
		return e
	}
	return s.formatCommit(t, out)
}
func formatPPr(raw []byte, attrs map[string]*string, p ParagraphFormatPatch) ([]byte, error) {
	updated, e := fmtOpening(raw, attrs)
	if e != nil {
		return nil, e
	}
	prefix := []byte(`<root xmlns:a="` + packaging.NSDrawingML + `">`)
	wrapped := append(append([]byte{}, prefix...), append(updated, []byte(`</root>`)...)...)
	d, e := losslessxml.Parse(wrapped)
	if e != nil {
		return nil, editRefusal("unsupported_structure", e.Error())
	}
	prop := d.Elements()[1]
	if prop.Name() != name(packaging.NSDrawingML, "pPr") {
		return nil, editRefusal("unsupported_structure", "paragraph property")
	}
	type part struct {
		key   string
		value *int64
		child string
	}
	wanted := []part{{"lnSpc", p.LinePercent, "spcPct"}, {"spcBef", p.SpaceBefore, "spcPts"}, {"spcAft", p.SpaceAfter, "spcPts"}}
	ordered := map[string]int{"lnSpc": 0, "spcBef": 1, "spcAft": 2, "buNone": 3, "buChar": 3, "buAutoNum": 3, "tabLst": 4, "defRPr": 5, "extLst": 6}
	previous := -1
	nodes := map[string]losslessxml.Element{}
	for _, n := range d.Elements() {
		parent, ok := n.Parent()
		if !ok || parent != prop {
			continue
		}
		o, ok := ordered[n.Name().Local]
		if !ok || n.Name().Space != packaging.NSDrawingML || o < previous {
			return nil, editRefusal("unsupported_structure", "paragraph property order")
		}
		previous = o
		if _, exists := nodes[n.Name().Local]; exists {
			return nil, editRefusal("unsupported_structure", "duplicate paragraph property")
		}
		nodes[n.Name().Local] = n
	}
	type replacement struct {
		a, b int
		v    []byte
	}
	edits := []replacement{}
	for _, w := range wanted {
		n, ok := nodes[w.key]
		if !ok || w.value == nil {
			continue
		}
		children := fmtChildren(d, n, name(packaging.NSDrawingML, w.child))
		if len(children) != 1 {
			return nil, editRefusal("unsupported_structure", "spacing choice")
		}
		for _, child := range d.Elements() {
			if parent, ok := child.Parent(); ok && parent == n && child != children[0] {
				return nil, editRefusal("unsupported_structure", "ambiguous spacing")
			}
			if parent, ok := child.Parent(); ok && parent == children[0] {
				return nil, editRefusal("unsupported_structure", "spacing leaf has descendants")
			}
		}
		a, b := n.SourceRange()
		edits = append(edits, replacement{a - len(prefix), b - len(prefix), []byte(fmt.Sprintf(`<a:%s><a:%s val="%d"/></a:%s>`, w.key, w.child, *w.value, w.key))})
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	for _, x := range edits {
		updated = append(append(append([]byte{}, updated[:x.a]...), x.v...), updated[x.b:]...)
	}
	d, e = losslessxml.Parse(append(append([]byte{}, prefix...), append(updated, []byte(`</root>`)...)...))
	if e != nil {
		return nil, editRefusal("unsupported_structure", e.Error())
	}
	prop = d.Elements()[1]
	add := []byte{}
	missing := []part{}
	for _, w := range wanted {
		if w.value != nil {
			if _, exists := nodes[w.key]; !exists {
				missing = append(missing, w)
				add = append(add, []byte(fmt.Sprintf(`<a:%s><a:%s val="%d"/></a:%s>`, w.key, w.child, *w.value, w.key))...)
			}
		}
	}
	if len(add) == 0 {
		return updated, nil
	}
	if bytes.HasSuffix(updated, []byte("/>")) {
		return append(append([]byte{}, updated[:len(updated)-2]...), append([]byte(">"), append(add, []byte("</a:pPr>")...)...)...), nil
	}
	// Insert each missing spacing choice in schema order, without moving any
	// existing child or altering its lexical bytes.
	for _, w := range missing {
		wrapped = append(append([]byte{}, prefix...), append(updated, []byte(`</root>`)...)...)
		d, e = losslessxml.Parse(wrapped)
		if e != nil {
			return nil, editRefusal("unsupported_structure", e.Error())
		}
		prop = d.Elements()[1]
		insert := fmtOpeningEnd(updated)
		if insert < 1 {
			return nil, editRefusal("unsupported_structure", "paragraph opening")
		}
		for _, node := range d.Elements() {
			parent, ok := node.Parent()
			if !ok || parent != prop {
				continue
			}
			if ordered[node.Name().Local] < ordered[w.key] {
				_, end := node.SourceRange()
				insert = end - len(prefix)
				continue
			}
			a, _ := node.SourceRange()
			insert = a - len(prefix)
			break
		}
		fragment := []byte(fmt.Sprintf(`<a:%s><a:%s val="%d"/></a:%s>`, w.key, w.child, *w.value, w.key))
		updated = append(append(append([]byte{}, updated[:insert]...), fragment...), updated[insert:]...)
	}
	return updated, nil
}
