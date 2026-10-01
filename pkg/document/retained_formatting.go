package document

import (
	"bytes"
	"encoding/xml"
	"sort"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// RetainedFormatTarget identifies one direct main-body paragraph or one run in it.
// Unlike the legacy value-model setters, edits splice only selected source XML.
type RetainedFormatTarget struct {
	session        *EditSession
	generation     uint64
	part, hash     string
	paragraph, run int
	doc            *losslessxml.Document
	property       losslessxml.Element
}

const retainedW = packaging.NSWordprocessingML

func retainedName(local string) xml.Name { return xml.Name{Space: retainedW, Local: local} }
func retainedChildren(d *losslessxml.Document, p losslessxml.Element, local string) []losslessxml.Element {
	var out []losslessxml.Element
	for _, n := range d.Elements() {
		if parent, ok := n.Parent(); ok && parent == p && n.Name() == retainedName(local) {
			out = append(out, n)
		}
	}
	return out
}
func retainedOne(d *losslessxml.Document, p losslessxml.Element, local string) (losslessxml.Element, error) {
	a := retainedChildren(d, p, local)
	if len(a) != 1 {
		return losslessxml.Element{}, editRefusal("unsupported_structure", "one direct "+local+" required")
	}
	return a[0], nil
}
func (s *EditSession) FindRetainedFormat(paragraph, run int) (*RetainedFormatTarget, error) {
	source, hash, e := s.pkg.Part(s.part)
	if e != nil {
		return nil, e
	}
	d, e := losslessxml.Parse(source)
	if e != nil {
		return nil, e
	}
	root := d.Elements()[0]
	body, e := retainedOne(d, root, "body")
	if e != nil {
		return nil, e
	}
	ps := retainedChildren(d, body, "p")
	if paragraph < 0 || paragraph >= len(ps) {
		return nil, editRefusal("missing_target", "paragraph index")
	}
	p := ps[paragraph]
	var owner losslessxml.Element
	name := "pPr"
	if run >= 0 {
		rs := retainedChildren(d, p, "r")
		if run >= len(rs) {
			return nil, editRefusal("missing_target", "run index")
		}
		owner = rs[run]
		name = "rPr"
		if len(retainedChildren(d, owner, "t")) != 1 {
			return nil, editRefusal("unsupported_structure", "plain run text required")
		}
	} else if run == -1 {
		owner = p
	} else {
		return nil, editRefusal("missing_target", "run index")
	}
	props := retainedChildren(d, owner, name)
	if len(props) != 1 {
		return nil, editRefusal("unsupported_structure", "unique existing property required")
	}
	if e = s.retainedProtect(d, p, owner, source); e != nil {
		return nil, e
	}
	// Refuse structural boundaries that would make a direct-property edit ambiguous.
	for _, n := range d.Elements() {
		if n != p && !retainedWithin(n, p) {
			continue
		}
		switch n.Name().Local {
		case "fldChar", "instrText", "hyperlink", "ins", "del", "moveFrom", "moveTo", "sdt", "drawing", "object", "pPrChange", "rPrChange", "bookmarkStart", "bookmarkEnd":
			return nil, editRefusal("unsupported_structure", "complex paragraph")
		}
	}
	return &RetainedFormatTarget{s, s.generation, s.part, hash, paragraph, run, d, props[0]}, nil
}
func (s *EditSession) retainedProtect(d *losslessxml.Document, p, owner losslessxml.Element, source []byte) error {
	g, e := s.pkg.Graph()
	if e != nil {
		return e
	}
	for _, edge := range g.Edges {
		if edge.Source != s.part || edge.Type != packaging.RelTypeSettings {
			continue
		}
		if edge.External {
			return editRefusal("unsupported_structure", "external settings")
		}
		blob, _, er := s.pkg.Part(edge.ResolvedPart)
		if er != nil {
			return er
		}
		settings, er := losslessxml.Parse(blob)
		if er != nil {
			return er
		}
		if settings.Elements()[0].Name() != retainedName("settings") {
			return editRefusal("unsupported_structure", "settings root")
		}
		for _, n := range settings.Elements() {
			if n.Name() == retainedName("trackRevisions") {
				return editRefusal("unsupported_structure", "track revisions")
			}
			if n.Name() == retainedName("documentProtection") {
				v := retainedAttr(n, "enforcement")
				switch v {
				case "0", "false", "off":
				case "1", "true", "on":
					return editRefusal("protected_operation", "document protection")
				default:
					return editRefusal("unsupported_structure", "unknown protection enforcement")
				}
			}
		}
	}
	// Review and range markup elsewhere in this main part is outside a plain
	// retained-format editor, even when the selected paragraph is ordinary.
	blocked := map[string]bool{"ins": true, "del": true, "moveFrom": true, "moveTo": true, "fldChar": true, "fldSimple": true, "instrText": true, "commentRangeStart": true, "commentRangeEnd": true, "bookmarkStart": true, "bookmarkEnd": true, "permStart": true, "permEnd": true}
	for _, n := range d.Elements() {
		if n.Name().Space == retainedW && (blocked[n.Name().Local] || strings.HasSuffix(n.Name().Local, "Change")) {
			return editRefusal("unsupported_structure", "review/field/range markup")
		}
	}
	// Require first and unique property child, a plain run and paragraph, and
	// reject comments/processing instructions in the edited subtree.
	for _, n := range d.Elements() {
		parent, ok := n.Parent()
		if !ok {
			continue
		}
		if (parent == p || parent == owner) && n.Name().Space != retainedW {
			return editRefusal("unsupported_structure", "foreign direct child")
		}
		if parent == p && n.Name() != retainedName("pPr") && n.Name() != retainedName("r") {
			return editRefusal("unsupported_structure", "paragraph wrapper")
		}
		if parent == owner && owner != p && n.Name() != retainedName("rPr") && n.Name() != retainedName("t") {
			return editRefusal("unsupported_structure", "non-text run content")
		}
	}
	ppr := retainedChildren(d, p, "pPr")
	if len(ppr) != 1 {
		return editRefusal("unsupported_structure", "paragraph properties")
	}
	ps, _ := p.SourceRange()
	pp, _ := ppr[0].SourceRange()
	firstPChild := -1
	for _, n := range d.Elements() {
		parent, ok := n.Parent()
		if ok && parent == p {
			start, _ := n.SourceRange()
			if firstPChild < 0 || start < firstPChild {
				firstPChild = start
			}
		}
	}
	if firstPChild != pp {
		return editRefusal("unsupported_structure", "pPr must be first")
	}
	if pp <= ps || bytes.Contains(source[ps:pp], []byte("<!--")) || bytes.Contains(source[ps:pp], []byte("<?")) {
		return editRefusal("unsupported_structure", "paragraph lexical barrier")
	}
	if owner != p {
		rs := retainedChildren(d, owner, "rPr")
		if len(rs) != 1 {
			return editRefusal("unsupported_structure", "run properties")
		}
		os, _ := owner.SourceRange()
		rp, _ := rs[0].SourceRange()
		firstRunChild := -1
		for _, n := range d.Elements() {
			parent, ok := n.Parent()
			if ok && parent == owner {
				start, _ := n.SourceRange()
				if firstRunChild < 0 || start < firstRunChild {
					firstRunChild = start
				}
			}
		}
		if firstRunChild != rp {
			return editRefusal("unsupported_structure", "rPr must be first")
		}
		if rp <= os || bytes.Contains(source[os:rp], []byte("<!--")) || bytes.Contains(source[os:rp], []byte("<?")) {
			return editRefusal("unsupported_structure", "run lexical barrier")
		}
	}
	a, b := p.SourceRange()
	if bytes.Contains(source[a:b], []byte("<!--")) || bytes.Contains(source[a:b], []byte("<?")) {
		return editRefusal("unsupported_structure", "paragraph comment or PI")
	}
	return nil
}
func retainedWithin(n, ancestor losslessxml.Element) bool {
	for p, ok := n.Parent(); ok; p, ok = p.Parent() {
		if p == ancestor {
			return true
		}
	}
	return false
}
func (s *EditSession) retainedCurrent(t *RetainedFormatTarget) error {
	if t == nil || t.session != s || t.generation != s.generation {
		return editRefusal("stale_target", "foreign or stale format target")
	}
	_, hash, e := s.pkg.Part(s.part)
	if e != nil {
		return e
	}
	if hash != t.hash {
		return editRefusal("stale_target", "format part changed")
	}
	return nil
}
func retainedAttr(n losslessxml.Element, key string) string {
	for _, a := range n.Attributes() {
		if a.Name == retainedName(key) {
			return a.Value
		}
	}
	return ""
}
func retainedOrder(d *losslessxml.Document, prop losslessxml.Element, order map[string]int) (map[string]losslessxml.Element, error) {
	found := map[string]losslessxml.Element{}
	last := -1
	for _, n := range d.Elements() {
		p, ok := n.Parent()
		if !ok || p != prop {
			continue
		}
		rank, valid := order[n.Name().Local]
		if n.Name().Space != retainedW || !valid || rank < last {
			return nil, editRefusal("unsupported_structure", "unsupported or disordered property child")
		}
		if _, exists := found[n.Name().Local]; exists {
			return nil, editRefusal("unsupported_structure", "duplicate property child")
		}
		found[n.Name().Local] = n
		last = rank
	}
	return found, nil
}

// A deliberately finite selected-property profile; unknown properties refuse rather than disappearing.
var retainedRunOrder = map[string]int{"rFonts": 1, "b": 2, "i": 3, "caps": 4, "smallCaps": 5, "strike": 6, "dstrike": 7, "vanish": 8, "color": 9, "sz": 10, "highlight": 11, "u": 12, "vertAlign": 13, "lang": 14}
var retainedParaOrder = map[string]int{"keepNext": 1, "keepLines": 2, "pageBreakBefore": 3, "widowControl": 4, "spacing": 5, "ind": 6, "jc": 7, "outlineLvl": 8}

func retainedOpeningEnd(raw []byte) int {
	q := byte(0)
	for i, c := range raw {
		if q != 0 {
			if c == q {
				q = 0
			}
			continue
		}
		if c == '\'' || c == '"' {
			q = c
			continue
		}
		if c == '>' {
			return i + 1
		}
	}
	return -1
}

// Attribute parser refuses prefixed lookalikes and duplicate raw names; it consumes quoted >.
func retainedOpening(raw []byte, changes map[string]*string) ([]byte, error) {
	end := retainedOpeningEnd(raw)
	if end < 0 {
		return nil, editRefusal("unsupported_structure", "missing opening")
	}
	type change struct {
		a, b  int
		value []byte
	}
	edits := []change{}
	seen := map[string]bool{}
	nameChar := func(c byte) bool {
		return c == '_' || c == ':' || c == '-' || c == '.' || c >= '0' && c <= '9' || c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z'
	}
	i := 1
	for i < end && nameChar(raw[i]) {
		i++
	}
	if i == 1 {
		return nil, editRefusal("unsupported_structure", "opening name")
	}
	for i < end {
		start := i
		for i < end && (raw[i] == ' ' || raw[i] == '\t' || raw[i] == '\n' || raw[i] == '\r') {
			i++
		}
		if i == end-1 && raw[i] == '>' || i == end-2 && raw[i] == '/' {
			break
		}
		if start == i {
			return nil, editRefusal("unsupported_structure", "attribute separator")
		}
		a := i
		for i < end && nameChar(raw[i]) {
			i++
		}
		if a == i {
			return nil, editRefusal("unsupported_structure", "attribute name")
		}
		key := string(raw[a:i])
		if seen[key] {
			return nil, editRefusal("unsupported_structure", "duplicate attribute")
		}
		seen[key] = true
		for i < end && (raw[i] == ' ' || raw[i] == '\t') {
			i++
		}
		if i >= end || raw[i] != '=' {
			return nil, editRefusal("unsupported_structure", "attribute assignment")
		}
		i++
		for i < end && (raw[i] == ' ' || raw[i] == '\t') {
			i++
		}
		if i >= end || (raw[i] != '\'' && raw[i] != '"') {
			return nil, editRefusal("unsupported_structure", "attribute quote")
		}
		q := raw[i]
		i++
		v := i
		for i < end && raw[i] != q {
			i++
		}
		if i >= end {
			return nil, editRefusal("unsupported_structure", "unterminated attribute")
		}
		vEnd := i
		i++
		if replacement, ok := changes[key]; ok {
			if replacement == nil {
				edits = append(edits, change{start, i, nil})
			} else if string(raw[v:vEnd]) != *replacement {
				edits = append(edits, change{v, vEnd, []byte(*replacement)})
			}
		}
	}
	for key, v := range changes {
		if seen[key] || v == nil {
			continue
		}
		if strings.Contains(key, ":") {
			return nil, editRefusal("unsupported_structure", "prefixed new attribute")
		}
		at := end - 1
		if raw[at-1] == '/' {
			at--
		}
		edits = append(edits, change{at, at, []byte(` ` + key + `="` + *v + `"`)})
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	out := bytes.Clone(raw)
	for _, v := range edits {
		out = append(append(bytes.Clone(out[:v.a]), v.value...), out[v.b:]...)
	}
	return out, nil
}
func retainedLeaf(key, value string) []byte { return []byte(`<w:` + key + ` w:val="` + value + `"/>`) }
func retainedRender(raw []byte, order map[string]int, updates map[string][]byte) ([]byte, error) {
	prefix := []byte(`<root xmlns:w="` + retainedW + `">`)
	wrapped := append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...)
	d, e := losslessxml.Parse(wrapped)
	if e != nil || len(d.Elements()) < 2 {
		return nil, editRefusal("unsupported_structure", "invalid property XML")
	}
	prop := d.Elements()[1]
	old, e := retainedOrder(d, prop, order)
	if e != nil {
		return nil, e
	}
	if e = retainedValidateLeaves(d, prop, old, raw, order); e != nil {
		return nil, e
	}
	type edit struct {
		a, b int
		data []byte
	}
	edits := []edit{}
	for key, v := range updates {
		if n, ok := old[key]; ok {
			a, b := n.SourceRange()
			edits = append(edits, edit{a - len(prefix), b - len(prefix), v})
		}
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	out := bytes.Clone(raw)
	for _, v := range edits {
		out = append(append(bytes.Clone(out[:v.a]), v.data...), out[v.b:]...)
	}
	// Sort absent insertions; map iteration cannot determine schema order.
	keys := make([]string, 0, len(updates))
	for key, v := range updates {
		if len(v) > 0 {
			if _, exists := old[key]; !exists {
				keys = append(keys, key)
			}
		}
	}
	sort.Slice(keys, func(i, j int) bool { return order[keys[i]] < order[keys[j]] })
	for _, key := range keys {
		v := updates[key]
		if len(v) == 0 {
			continue
		}
		if _, exists := old[key]; exists {
			continue
		}
		wrapped = append(append(bytes.Clone(prefix), out...), []byte(`</root>`)...)
		d, e = losslessxml.Parse(wrapped)
		if e != nil {
			return nil, editRefusal("unsupported_structure", "property insert parse")
		}
		prop = d.Elements()[1]
		at := retainedOpeningEnd(out)
		if at < 0 {
			return nil, editRefusal("unsupported_structure", "property opening")
		}
		for _, n := range d.Elements() {
			p, ok := n.Parent()
			if ok && p == prop && order[n.Name().Local] < order[key] {
				_, b := n.SourceRange()
				if b-len(prefix) > at {
					at = b - len(prefix)
				}
			}
		}
		if bytes.HasSuffix(out, []byte("/>")) {
			opening := retainedOpeningEnd(out)
			nameEnd := 1
			for nameEnd < len(out) && out[nameEnd] != ' ' && out[nameEnd] != '\t' && out[nameEnd] != '\r' && out[nameEnd] != '\n' && out[nameEnd] != '/' && out[nameEnd] != '>' {
				nameEnd++
			}
			if nameEnd == 1 {
				return nil, editRefusal("unsupported_structure", "property QName")
			}
			out = append(append(bytes.Clone(out[:opening-2]), '>'), []byte("</"+string(out[1:nameEnd])+">")...)
			at = opening - 1
		}
		out = append(append(bytes.Clone(out[:at]), v...), out[at:]...)
	}
	return out, nil
}
func retainedValidateLeaves(d *losslessxml.Document, prop losslessxml.Element, children map[string]losslessxml.Element, raw []byte, order map[string]int) error {
	if prop.Name() != retainedName("rPr") && prop.Name() != retainedName("pPr") {
		return editRefusal("unsupported_structure", "property QName")
	}
	for _, attr := range prop.Attributes() {
		return editRefusal("unsupported_structure", "property attributes unsupported: "+attr.Name.Local)
	}
	if bytes.Contains(raw, []byte("<!--")) || bytes.Contains(raw, []byte("<?")) {
		return editRefusal("unsupported_structure", "property lexical barrier")
	}
	for _, n := range children {
		if n.Name().Space != retainedW {
			return editRefusal("unsupported_structure", "foreign property child")
		}
		allowed := map[string]map[string]bool{"rFonts": {"ascii": true, "hAnsi": true}, "spacing": {"before": true, "after": true, "line": true, "lineRule": true}, "ind": {"left": true, "right": true, "hanging": true}}
		for _, a := range n.Attributes() {
			if a.Name.Space != retainedW {
				return editRefusal("unsupported_structure", "foreign leaf attribute")
			}
			if a.Name.Local != "val" && !allowed[n.Name().Local][a.Name.Local] {
				return editRefusal("unsupported_structure", "unknown leaf attribute")
			}
			if a.Name.Local == "val" && len(allowed[n.Name().Local]) > 0 {
				return editRefusal("unsupported_structure", "unexpected leaf val")
			}
		}
		for _, child := range d.Elements() {
			if p, ok := child.Parent(); ok && p == n {
				return editRefusal("unsupported_structure", "property leaf has descendants")
			}
		}
		frag := n.Raw()
		if bytes.Contains(frag, []byte("<!--")) || bytes.Contains(frag, []byte("<?")) {
			return editRefusal("unsupported_structure", "property lexical barrier")
		}
	}
	_ = order
	return nil
}
func retainedInt(v any, min, max int64) (string, error) {
	n, ok := v.(int64)
	if !ok || n < min || n > max {
		return "", editRefusal("invalid-format", "integer bounds")
	}
	return strconv.FormatInt(n, 10), nil
}
func retainedBoolean(v any) (string, error) {
	b, ok := v.(bool)
	if !ok {
		return "", editRefusal("invalid-format", "typed boolean")
	}
	if b {
		return "1", nil
	}
	return "0", nil
}
func retainedHex(v any) (string, error) {
	s, ok := v.(string)
	if !ok || len(s) != 6 {
		return "", editRefusal("invalid-format", "hex colour")
	}
	for _, r := range s {
		if r < '0' || r > '9' && r < 'A' || r > 'F' {
			return "", editRefusal("invalid-format", "uppercase hex colour")
		}
	}
	return s, nil
}

// SetRetainedFormat applies a finite direct-property patch to the selected run
// (run >= 0) or paragraph (run == -1), without serialising a legacy model.
func (s *EditSession) SetRetainedFormat(t *RetainedFormatTarget, patch map[string]any) error {
	if e := s.retainedCurrent(t); e != nil {
		return e
	}
	if len(patch) == 0 {
		return editRefusal("invalid-format", "empty format patch")
	}
	order := retainedRunOrder
	if t.run < 0 {
		order = retainedParaOrder
	}
	source, _, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	a, b := t.property.SourceRange()
	raw := bytes.Clone(source[a:b])
	updates := map[string][]byte{}
	// Use a local parser so the exact selected property child values and topology
	// are checked before forming any replacement bytes.
	prefix := []byte(`<root xmlns:w="` + retainedW + `">`)
	d, e := losslessxml.Parse(append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...))
	if e != nil {
		return editRefusal("unsupported_structure", "property parse")
	}
	children, e := retainedOrder(d, d.Elements()[1], order)
	if e != nil {
		return e
	}
	for key, value := range patch {
		if t.run >= 0 {
			switch key {
			case "strike", "caps", "smallCaps", "vanish":
				v, er := retainedBoolean(value)
				if er != nil {
					return er
				}
				tag := key
				if key == "caps" {
					tag = "caps"
				}
				if key == "smallCaps" {
					tag = "smallCaps"
				}
				updates[tag] = retainedLeaf(tag, v)
			case "underline", "highlight", "verticalAlign":
				v, ok := value.(string)
				if !ok {
					return editRefusal("invalid-format", "string choice")
				}
				allowed := map[string]map[string]bool{"underline": {"double": true}, "highlight": {"yellow": true}, "verticalAlign": {"superscript": true, "subscript": true}}
				if !allowed[key][v] {
					return editRefusal("invalid-format", "enum choice")
				}
				tag := map[string]string{"underline": "u", "highlight": "highlight", "verticalAlign": "vertAlign"}[key]
				updates[tag] = retainedLeaf(tag, v)
			case "color":
				v, er := retainedHex(value)
				if er != nil {
					return er
				}
				updates["color"] = retainedLeaf("color", v)
			case "sizeHalfPoints":
				v, er := retainedInt(value, 1, 800)
				if er != nil {
					return er
				}
				updates["sz"] = retainedLeaf("sz", v)
			case "ascii", "hAnsi": // validated jointly below
			case "removeDirect":
				if value != true {
					return editRefusal("invalid-format", "removeDirect must be true")
				}
				for _, k := range []string{"rFonts", "color", "sz", "highlight", "u", "vertAlign"} {
					updates[k] = nil
				}
			default:
				return editRefusal("invalid-format", "unknown run property")
			}
		} else {
			switch key {
			case "jc":
				v, ok := value.(string)
				if !ok || v != "both" {
					return editRefusal("invalid-format", "alignment")
				}
				updates["jc"] = retainedLeaf("jc", v)
			case "before", "after", "line", "lineRule": // spacing grouped below
			case "left", "hanging": // indentation grouped below
			case "keepLines", "keepNext", "pageBreakBefore":
				v, er := retainedBoolean(value)
				if er != nil {
					return er
				}
				updates[key] = retainedLeaf(key, v)
			case "outlineLvl":
				v, er := retainedInt(value, 0, 9)
				if er != nil {
					return er
				}
				updates["outlineLvl"] = retainedLeaf("outlineLvl", v)
			default:
				return editRefusal("invalid-format", "unknown paragraph property")
			}
		}
	}
	if t.run >= 0 {
		ascii, aok := patch["ascii"]
		hansi, hok := patch["hAnsi"]
		if aok != hok {
			return editRefusal("invalid-format", "paired fonts required")
		}
		if aok {
			a, aok := ascii.(string)
			h, hok := hansi.(string)
			if !aok || !hok || a != h || a == "" || strings.TrimSpace(a) != a || len([]rune(a)) > 128 || strings.ContainsAny(a, "<>\"'\r\n\t&") {
				return editRefusal("invalid-format", "font name")
			}
			n := children["rFonts"]
			if n.Ordinal() < 0 {
				return editRefusal("unsupported_structure", "paired font child required")
			}
			fontChanges := map[string]*string{"w:ascii": &a, "w:hAnsi": &a}
			updated, er := retainedOpening([]byte(n.Raw()), fontChanges)
			if er != nil {
				return er
			}
			updates["rFonts"] = updated
		}
	} else {
		if _, ok := patch["before"]; ok || patch["after"] != nil || patch["line"] != nil || patch["lineRule"] != nil {
			n, ok := children["spacing"]
			if !ok {
				return editRefusal("unsupported_structure", "spacing child required")
			}
			rawSpacing := []byte(n.Raw())
			changes := map[string]*string{}
			for _, k := range []string{"before", "after", "line"} {
				if val, ok := patch[k]; ok {
					min := int64(0)
					if k == "line" {
						min = 1
					}
					v, er := retainedInt(val, min, 31680)
					if er != nil {
						return er
					}
					changes["w:"+k] = &v
				}
			}
			if val, ok := patch["lineRule"]; ok {
				v, ok := val.(string)
				if !ok || v != "auto" && v != "exact" && v != "atLeast" {
					return editRefusal("invalid-format", "line rule")
				}
				changes["w:lineRule"] = &v
			}
			updated, er := retainedOpening(rawSpacing, changes)
			if er != nil {
				return er
			}
			updates["spacing"] = updated
		}
		if _, ok := patch["left"]; ok || patch["hanging"] != nil {
			n, ok := children["ind"]
			if !ok {
				return editRefusal("unsupported_structure", "indent child required")
			}
			changes := map[string]*string{}
			for _, k := range []string{"left", "hanging"} {
				if val, ok := patch[k]; ok {
					v, er := retainedInt(val, 0, 31680)
					if er != nil {
						return er
					}
					changes["w:"+k] = &v
				}
			}
			updated, er := retainedOpening([]byte(n.Raw()), changes)
			if er != nil {
				return er
			}
			updates["ind"] = updated
		}
	}
	updated, e := retainedRender(raw, order, updates)
	if e != nil {
		return e
	}
	out := append(append(bytes.Clone(source[:a]), updated...), source[b:]...)
	if _, e = losslessxml.Parse(out); e != nil {
		return editRefusal("unsupported_structure", "result parse")
	}
	if bytes.Equal(source, out) {
		return nil
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: t.part, ExpectedSHA256: t.hash, Data: out}}); e != nil {
		return e
	}
	s.generation++
	return nil
}
