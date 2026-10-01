package document

import (
	"bytes"
	"encoding/xml"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// RetainedTableTarget identifies one direct main-body table and a physical cell.
type RetainedTableTarget struct {
	session          *EditSession
	generation       uint64
	part, hash       string
	doc              *losslessxml.Document
	table, row, cell losslessxml.Element
}

func wordTableRefusal(kind, detail string) error { return editRefusal(kind, detail) }
func wordTableChildren(d *losslessxml.Document, p losslessxml.Element, key string) []losslessxml.Element {
	return retainedChildren(d, p, key)
}
func wordTableOne(d *losslessxml.Document, p losslessxml.Element, key string) (losslessxml.Element, error) {
	return retainedOne(d, p, key)
}
func wordTableInt(v any, min, max int64) (string, error) {
	var n int64
	switch x := v.(type) {
	case int:
		n = int64(x)
	case int64:
		n = x
	case float64:
		if x != float64(int64(x)) {
			return "", wordTableRefusal("invalid-table-properties", "integer required")
		}
		n = int64(x)
	default:
		return "", wordTableRefusal("invalid-table-properties", "integer required")
	}
	if n < min || n > max {
		return "", wordTableRefusal("invalid-table-properties", "integer bounds")
	}
	return strconv.FormatInt(n, 10), nil
}
func wordTableBool(v any) (string, error) {
	b, ok := v.(bool)
	if !ok {
		return "", wordTableRefusal("invalid-table-properties", "boolean required")
	}
	if b {
		return "1", nil
	}
	return "0", nil
}
func wordTableString(v any, choices ...string) (string, error) {
	s, ok := v.(string)
	if !ok {
		return "", wordTableRefusal("invalid-table-properties", "string required")
	}
	for _, c := range choices {
		if c == s {
			return s, nil
		}
	}
	return "", wordTableRefusal("invalid-table-properties", "unsupported choice")
}
func wordTableHex(v any) (string, error) {
	s, ok := v.(string)
	if !ok || len(s) != 6 {
		return "", wordTableRefusal("invalid-table-properties", "colour")
	}
	for _, c := range s {
		if c < '0' || c > '9' && c < 'A' || c > 'F' {
			return "", wordTableRefusal("invalid-table-properties", "colour")
		}
	}
	return s, nil
}
func wordTableAttrs(n losslessxml.Element, attrs ...string) error {
	allowed := map[string]bool{}
	for _, x := range attrs {
		allowed[x] = true
	}
	seen := map[string]bool{}
	for _, a := range n.Attributes() {
		if a.Name.Space != retainedW || !allowed[a.Name.Local] || seen[a.Name.Local] {
			return wordTableRefusal("unsupported_structure", "foreign/duplicate attribute")
		}
		seen[a.Name.Local] = true
	}
	return nil
}
func wordTableRaw(source []byte, n losslessxml.Element) ([]byte, error) {
	a, b := n.SourceRange()
	if a < 0 || b > len(source) || a >= b {
		return nil, wordTableRefusal("unsupported_structure", "invalid source range")
	}
	return bytes.Clone(source[a:b]), nil
}
func wordTableWrap(raw []byte) (*losslessxml.Document, losslessxml.Element, int, error) {
	prefix := []byte(`<root xmlns:w="` + retainedW + `">`)
	d, e := losslessxml.Parse(append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...))
	if e != nil || len(d.Elements()) < 2 {
		return nil, losslessxml.Element{}, 0, wordTableRefusal("unsupported_structure", "property XML")
	}
	return d, d.Elements()[1], len(prefix), nil
}
func wordTablePatch(raw []byte, n losslessxml.Element, repl []byte, off int) ([]byte, error) {
	a, b := n.SourceRange()
	a -= off
	b -= off
	if a < 0 || a > b || b > len(raw) {
		return nil, wordTableRefusal("unsupported_structure", "nested source range")
	}
	return append(append(bytes.Clone(raw[:a]), repl...), raw[b:]...), nil
}
func wordTableAllowed(d *losslessxml.Document, owner losslessxml.Element, order map[string]int) (map[string]losslessxml.Element, error) {
	prev := 0
	seen := map[string]losslessxml.Element{}
	for _, n := range d.Elements() {
		p, ok := n.Parent()
		if !ok || p != owner {
			continue
		}
		rank := order[n.Name().Local]
		if n.Name().Space != retainedW || rank == 0 || rank < prev || seen[n.Name().Local].Ordinal() >= 0 {
			return nil, wordTableRefusal("unsupported_structure", "property order/duplicate")
		}
		prev = rank
		seen[n.Name().Local] = n
	}
	return seen, nil
}

var wordCellOrder = map[string]int{"tcW": 1, "tcBorders": 2, "shd": 3, "noWrap": 4, "tcMar": 5, "textDirection": 6, "tcFitText": 7, "vAlign": 8, "hideMark": 9}
var wordBorderOrder = map[string]int{"top": 1, "left": 2, "bottom": 3, "right": 4}
var wordMarginOrder = map[string]int{"top": 1, "left": 2, "bottom": 3, "right": 4}
var wordRowOrder = map[string]int{"cantSplit": 1, "trHeight": 2, "tblHeader": 3}
var wordTableOrder = map[string]int{"tblStyle": 1, "tblW": 2, "jc": 3, "tblInd": 4, "tblLook": 5}

func wordTableLeaf(d *losslessxml.Document, n losslessxml.Element, attrs ...string) error {
	if e := wordTableAttrs(n, attrs...); e != nil {
		return e
	}
	for _, child := range d.Elements() {
		p, ok := child.Parent()
		if ok && p == n {
			return wordTableRefusal("unsupported_structure", "property child descendant")
		}
	}
	return nil
}
func wordTableScalar(n losslessxml.Element, k string, min, max int64) error {
	raw := retainedAttr(n, k)
	if raw == "" {
		return wordTableRefusal("unsupported_structure", "missing "+k)
	}
	i, e := strconv.ParseInt(raw, 10, 64)
	if e != nil || i < min || i > max {
		return wordTableRefusal("unsupported_structure", "invalid original "+k)
	}
	return nil
}
func wordTableEnum(n losslessxml.Element, k string, choices ...string) error {
	raw := retainedAttr(n, k)
	for _, v := range choices {
		if raw == v {
			return nil
		}
	}
	return wordTableRefusal("unsupported_structure", "invalid original "+k)
}
func wordTableValidate(d *losslessxml.Document, tbl, row, cell losslessxml.Element) error {
	tc, e := wordTableOne(d, cell, "tcPr")
	if e != nil {
		return e
	}
	tp, e := wordTableOne(d, tbl, "tblPr")
	if e != nil {
		return e
	}
	rp, e := wordTableOne(d, row, "trPr")
	if e != nil {
		return e
	}
	for _, item := range []struct {
		n     losslessxml.Element
		order map[string]int
	}{{tc, wordCellOrder}, {tp, wordTableOrder}, {rp, wordRowOrder}} {
		if len(item.n.Attributes()) != 0 {
			return wordTableRefusal("unsupported_structure", "property attributes")
		}
		if _, e = wordTableAllowed(d, item.n, item.order); e != nil {
			return e
		}
	}
	cellProps, _ := wordTableAllowed(d, tc, wordCellOrder)
	rowProps, _ := wordTableAllowed(d, rp, wordRowOrder)
	tableProps, _ := wordTableAllowed(d, tp, wordTableOrder)
	w := cellProps["tcW"]
	if w.Ordinal() < 0 || wordTableLeaf(d, w, "w", "type") != nil || wordTableEnum(w, "type", "dxa") != nil || wordTableScalar(w, "w", 0, 31680) != nil {
		return wordTableRefusal("unsupported_structure", "cell preferred width")
	}
	for _, item := range []struct {
		name  string
		attrs []string
	}{{"shd", []string{"val", "color", "fill"}}, {"textDirection", []string{"val"}}, {"vAlign", []string{"val"}}, {"noWrap", []string{"val"}}, {"tcFitText", []string{"val"}}, {"hideMark", []string{"val"}}} {
		n := cellProps[item.name]
		if n.Ordinal() < 0 || wordTableLeaf(d, n, item.attrs...) != nil {
			return wordTableRefusal("unsupported_structure", "cell leaf "+item.name)
		}
	}
	for _, k := range []string{"noWrap", "tcFitText", "hideMark"} {
		if e = wordTableEnum(cellProps[k], "val", "0", "1"); e != nil {
			return e
		}
	}
	if e = wordTableEnum(cellProps["textDirection"], "val", "lrTb", "tbRl"); e != nil {
		return e
	}
	if e = wordTableEnum(cellProps["vAlign"], "val", "top", "center"); e != nil {
		return e
	}
	shd := cellProps["shd"]
	if e = wordTableEnum(shd, "val", "clear"); e != nil {
		return e
	}
	if e = wordTableEnum(shd, "color", "auto"); e != nil {
		return e
	}
	if _, e = wordTableHex(retainedAttr(shd, "fill")); e != nil {
		return wordTableRefusal("unsupported_structure", "original shading")
	}
	for _, parent := range []struct {
		name  string
		order map[string]int
		attrs []string
	}{{"tcBorders", wordBorderOrder, []string{"val", "sz", "color"}}, {"tcMar", wordMarginOrder, []string{"w", "type"}}} {
		n := cellProps[parent.name]
		if n.Ordinal() < 0 || len(n.Attributes()) != 0 {
			return wordTableRefusal("unsupported_structure", "missing "+parent.name)
		}
		children, er := wordTableAllowed(d, n, parent.order)
		if er != nil {
			return er
		}
		for _, side := range []string{"top", "left", "bottom", "right"} {
			child := children[side]
			if child.Ordinal() < 0 || wordTableLeaf(d, child, parent.attrs...) != nil {
				return wordTableRefusal("unsupported_structure", "missing side "+side)
			}
			if parent.name == "tcBorders" {
				if wordTableEnum(child, "val", "single") != nil || wordTableScalar(child, "sz", 2, 96) != nil {
					return wordTableRefusal("unsupported_structure", "border value")
				}
				if _, er = wordTableHex(retainedAttr(child, "color")); er != nil {
					return wordTableRefusal("unsupported_structure", "border colour")
				}
			} else if wordTableEnum(child, "type", "dxa") != nil || wordTableScalar(child, "w", 0, 31680) != nil {
				return wordTableRefusal("unsupported_structure", "margin value")
			}
		}
	}
	for _, k := range []string{"cantSplit", "trHeight", "tblHeader"} {
		if rowProps[k].Ordinal() < 0 {
			return wordTableRefusal("unsupported_structure", "missing row property")
		}
	}
	for _, k := range []string{"cantSplit", "tblHeader"} {
		if e = wordTableLeaf(d, rowProps[k], "val"); e != nil {
			return e
		}
		if e = wordTableEnum(rowProps[k], "val", "0", "1"); e != nil {
			return e
		}
	}
	height := rowProps["trHeight"]
	if e = wordTableLeaf(d, height, "val", "hRule"); e != nil {
		return e
	}
	if e = wordTableScalar(height, "val", 1, 31680); e != nil {
		return e
	}
	if e = wordTableEnum(height, "hRule", "atLeast", "exact"); e != nil {
		return e
	}
	if tableProps["jc"].Ordinal() < 0 || tableProps["tblInd"].Ordinal() < 0 {
		return wordTableRefusal("unsupported_structure", "missing table property")
	}
	if e = wordTableLeaf(d, tableProps["jc"], "val"); e != nil {
		return e
	}
	if e = wordTableEnum(tableProps["jc"], "val", "left", "center"); e != nil {
		return e
	}
	if e = wordTableLeaf(d, tableProps["tblInd"], "w", "type"); e != nil {
		return e
	}
	if e = wordTableEnum(tableProps["tblInd"], "type", "dxa"); e != nil {
		return e
	}
	return wordTableScalar(tableProps["tblInd"], "w", 0, 31680)
}

func (s *EditSession) FindRetainedTable(tableIndex, rowIndex, cellIndex int) (*RetainedTableTarget, error) {
	if tableIndex < 0 || rowIndex < 0 || cellIndex < 0 {
		return nil, wordTableRefusal("missing_target", "table indexes")
	}
	source, hash, e := s.pkg.Part(s.part)
	if e != nil {
		return nil, e
	}
	d, e := losslessxml.Parse(source)
	if e != nil {
		return nil, e
	}
	nodes := d.Elements()
	if len(nodes) == 0 || nodes[0].Name() != retainedName("document") {
		return nil, wordTableRefusal("unsupported_structure", "document root")
	}
	body, e := wordTableOne(d, nodes[0], "body")
	if e != nil {
		return nil, e
	}
	tables := wordTableChildren(d, body, "tbl")
	if tableIndex >= len(tables) {
		return nil, wordTableRefusal("missing_target", "table index")
	}
	tbl := tables[tableIndex]
	rows := wordTableChildren(d, tbl, "tr")
	if rowIndex >= len(rows) {
		return nil, wordTableRefusal("missing_target", "row index")
	}
	row := rows[rowIndex]
	cells := wordTableChildren(d, row, "tc")
	if cellIndex >= len(cells) {
		return nil, wordTableRefusal("missing_target", "cell index")
	}
	cell := cells[cellIndex]
	if e = s.wordTableProtection(d); e != nil {
		return nil, e
	}
	if e = wordTableAdmit(d, tbl); e != nil {
		return nil, e
	}
	for _, n := range []losslessxml.Element{tbl, row, cell} {
		raw, er := wordTableRaw(source, n)
		if er != nil {
			return nil, er
		}
		if bytes.Contains(raw, []byte("<!--")) || bytes.Contains(raw, []byte("<?")) {
			return nil, wordTableRefusal("unsupported_structure", "lexical table barrier")
		}
	}
	if e = wordTableValidate(d, tbl, row, cell); e != nil {
		return nil, e
	}
	return &RetainedTableTarget{session: s, generation: s.generation, part: s.part, hash: hash, doc: d, table: tbl, row: row, cell: cell}, nil
}
func (s *EditSession) wordTableCurrent(t *RetainedTableTarget) error {
	if t == nil || t.session != s || t.generation != s.generation {
		return wordTableRefusal("stale_target", "foreign or stale table")
	}
	_, hash, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	if hash != t.hash {
		return wordTableRefusal("stale_target", "table part changed")
	}
	return nil
}
func wordTableRender(raw []byte, owner string, order map[string]int, updates map[string][]byte) ([]byte, error) {
	d, n, off, e := wordTableWrap(raw)
	if e != nil {
		return nil, e
	}
	if n.Name() != (xml.Name{Space: retainedW, Local: owner}) {
		return nil, wordTableRefusal("unsupported_structure", "wrong owner")
	}
	children, e := wordTableAllowed(d, n, order)
	if e != nil {
		return nil, e
	}
	type edit struct {
		a, b int
		v    []byte
	}
	edits := []edit{}
	for key, v := range updates {
		if order[key] == 0 {
			return nil, wordTableRefusal("invalid-table-properties", "unsupported property")
		}
		old, ok := children[key]
		if !ok {
			if v == nil {
				continue
			}
			return nil, wordTableRefusal("unsupported_structure", "missing property")
		}
		a, b := old.SourceRange()
		edits = append(edits, edit{a - off, b - off, v})
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	out := bytes.Clone(raw)
	for _, item := range edits {
		if item.a < 0 || item.b > len(out) {
			return nil, wordTableRefusal("unsupported_structure", "property range")
		}
		out = append(append(bytes.Clone(out[:item.a]), item.v...), out[item.b:]...)
	}
	return out, nil
}
func wordTableUpdateLeaf(d *losslessxml.Document, n losslessxml.Element, source []byte, changes map[string]*string) ([]byte, error) {
	raw, e := wordTableRaw(source, n)
	if e != nil {
		return nil, e
	}
	return retainedOpening(raw, changes)
}
func wordTableLeafChange(d *losslessxml.Document, p losslessxml.Element, key string, order map[string]int, source []byte, changes map[string]*string) ([]byte, error) {
	children, e := wordTableAllowed(d, p, order)
	if e != nil {
		return nil, e
	}
	n := children[key]
	if n.Ordinal() < 0 {
		return nil, wordTableRefusal("unsupported_structure", "missing "+key)
	}
	return wordTableUpdateLeaf(d, n, source, changes)
}

// SetRetainedTable splices only the admitted table properties after complete preflight.
func (s *EditSession) SetRetainedTable(t *RetainedTableTarget, patch map[string]any) error {
	if e := s.wordTableCurrent(t); e != nil {
		return e
	}
	if len(patch) == 0 {
		return wordTableRefusal("invalid-table-properties", "empty patch")
	}
	valid := map[string]bool{"shading": true, "verticalAlign": true, "textDirection": true, "margins": true, "border": true, "noWrap": true, "tcFitText": true, "width": true, "tblHeader": true, "cantSplit": true, "height": true, "heightRule": true, "alignment": true, "indent": true, "hideMark": true}
	for key := range patch {
		if !valid[key] {
			return wordTableRefusal("invalid-table-properties", "unknown patch key")
		}
	}
	source, _, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	d := t.doc
	if e = wordTableAdmit(d, t.table); e != nil {
		return e
	}
	if e = wordTableValidate(d, t.table, t.row, t.cell); e != nil {
		return e
	}
	tc, _ := wordTableOne(d, t.cell, "tcPr")
	tr, _ := wordTableOne(d, t.row, "trPr")
	tbl, _ := wordTableOne(d, t.table, "tblPr")
	cellUpdates := map[string][]byte{}
	rowUpdates := map[string][]byte{}
	tableUpdates := map[string][]byte{}
	cellChildren, _ := wordTableAllowed(d, tc, wordCellOrder)
	rowChildren, _ := wordTableAllowed(d, tr, wordRowOrder)
	tableChildren, _ := wordTableAllowed(d, tbl, wordTableOrder)
	if v, ok := patch["shading"]; ok {
		if v == nil {
			cellUpdates["shd"] = nil
		} else {
			color, er := wordTableHex(v)
			if er != nil {
				return er
			}
			raw, er := wordTableUpdateLeaf(d, cellChildren["shd"], source, map[string]*string{"w:fill": &color})
			if er != nil {
				return er
			}
			cellUpdates["shd"] = raw
		}
	}
	for _, x := range []struct {
		patch, key string
		choices    []string
	}{{"verticalAlign", "vAlign", []string{"top", "center"}}, {"textDirection", "textDirection", []string{"lrTb", "tbRl"}}} {
		if v, ok := patch[x.patch]; ok {
			val, er := wordTableString(v, x.choices...)
			if er != nil {
				return er
			}
			raw, er := wordTableUpdateLeaf(d, cellChildren[x.key], source, map[string]*string{"w:val": &val})
			if er != nil {
				return er
			}
			cellUpdates[x.key] = raw
		}
	}
	if v, ok := patch["margins"]; ok {
		m, yes := v.(map[string]any)
		if !yes || len(m) != 4 {
			return wordTableRefusal("invalid-table-properties", "four margins required")
		}
		parent := cellChildren["tcMar"]
		sideNodes, _ := wordTableAllowed(d, parent, wordMarginOrder)
		raw, er := wordTableRaw(source, parent)
		if er != nil {
			return er
		}
		type change struct {
			a, b int
			v    []byte
		}
		edits := []change{}
		a, _ := parent.SourceRange()
		for _, side := range []string{"top", "left", "bottom", "right"} {
			val, er := wordTableInt(m[side], 0, 31680)
			if er != nil {
				return er
			}
			n := sideNodes[side]
			updated, er := wordTableUpdateLeaf(d, n, source, map[string]*string{"w:w": &val})
			if er != nil {
				return er
			}
			x, y := n.SourceRange()
			edits = append(edits, change{x - a, y - a, updated})
		}
		for i := 0; i < len(edits); i++ {
			for j := i + 1; j < len(edits); j++ {
				if edits[i].a < edits[j].a {
					edits[i], edits[j] = edits[j], edits[i]
				}
			}
		}
		for _, ed := range edits {
			raw = append(append(bytes.Clone(raw[:ed.a]), ed.v...), raw[ed.b:]...)
		}
		cellUpdates["tcMar"] = raw
	}
	if v, ok := patch["border"]; ok {
		b, yes := v.(map[string]any)
		if !yes {
			return wordTableRefusal("invalid-table-properties", "border object")
		}
		side, er := wordTableString(b["side"], "top", "left", "bottom", "right")
		if er != nil {
			return er
		}
		for k := range b {
			if k != "side" && k != "style" && k != "size" && k != "color" && k != "remove" {
				return wordTableRefusal("invalid-table-properties", "unknown border key")
			}
		}
		parent := cellChildren["tcBorders"]
		sides, _ := wordTableAllowed(d, parent, wordBorderOrder)
		n := sides[side]
		raw, er := wordTableRaw(source, parent)
		if er != nil {
			return er
		}
		a, _ := parent.SourceRange()
		x, y := n.SourceRange()
		if remove, exists := b["remove"]; exists {
			if remove != true || len(b) != 2 {
				return wordTableRefusal("invalid-table-properties", "remove border")
			}
			raw = append(bytes.Clone(raw[:x-a]), raw[y-a:]...)
		} else {
			changes := map[string]*string{}
			if z, found := b["style"]; found {
				v, er := wordTableString(z, "single")
				if er != nil {
					return er
				}
				changes["w:val"] = &v
			}
			if z, found := b["size"]; found {
				v, er := wordTableInt(z, 2, 96)
				if er != nil {
					return er
				}
				changes["w:sz"] = &v
			}
			if z, found := b["color"]; found {
				v, er := wordTableHex(z)
				if er != nil {
					return er
				}
				changes["w:color"] = &v
			}
			if len(changes) == 0 {
				return wordTableRefusal("invalid-table-properties", "empty border")
			}
			updated, er := wordTableUpdateLeaf(d, n, source, changes)
			if er != nil {
				return er
			}
			raw = append(append(bytes.Clone(raw[:x-a]), updated...), raw[y-a:]...)
		}
		cellUpdates["tcBorders"] = raw
	}
	for _, key := range []string{"noWrap", "tcFitText", "hideMark"} {
		if x, ok := patch[key]; ok {
			v, er := wordTableBool(x)
			if er != nil {
				return er
			}
			raw, er := wordTableUpdateLeaf(d, cellChildren[key], source, map[string]*string{"w:val": &v})
			if er != nil {
				return er
			}
			cellUpdates[key] = raw
		}
	}
	if x, ok := patch["width"]; ok {
		v, er := wordTableInt(x, 0, 31680)
		if er != nil {
			return er
		}
		raw, er := wordTableUpdateLeaf(d, cellChildren["tcW"], source, map[string]*string{"w:w": &v})
		if er != nil {
			return er
		}
		cellUpdates["tcW"] = raw
	}
	for _, key := range []string{"tblHeader", "cantSplit"} {
		if x, ok := patch[key]; ok {
			v, er := wordTableBool(x)
			if er != nil {
				return er
			}
			raw, er := wordTableUpdateLeaf(d, rowChildren[key], source, map[string]*string{"w:val": &v})
			if er != nil {
				return er
			}
			rowUpdates[key] = raw
		}
	}
	if _, ok := patch["height"]; ok || patch["heightRule"] != nil {
		v, er := wordTableInt(patch["height"], 1, 31680)
		if er != nil {
			return er
		}
		rule, er := wordTableString(patch["heightRule"], "atLeast", "exact")
		if er != nil {
			return er
		}
		raw, er := wordTableUpdateLeaf(d, rowChildren["trHeight"], source, map[string]*string{"w:val": &v, "w:hRule": &rule})
		if er != nil {
			return er
		}
		rowUpdates["trHeight"] = raw
	}
	if x, ok := patch["alignment"]; ok {
		v, er := wordTableString(x, "left", "center")
		if er != nil {
			return er
		}
		raw, er := wordTableUpdateLeaf(d, tableChildren["jc"], source, map[string]*string{"w:val": &v})
		if er != nil {
			return er
		}
		tableUpdates["jc"] = raw
	}
	if x, ok := patch["indent"]; ok {
		v, er := wordTableInt(x, 0, 31680)
		if er != nil {
			return er
		}
		raw, er := wordTableUpdateLeaf(d, tableChildren["tblInd"], source, map[string]*string{"w:w": &v})
		if er != nil {
			return er
		}
		tableUpdates["tblInd"] = raw
	}
	type edit struct {
		a, b int
		v    []byte
	}
	edits := []edit{}
	for _, item := range []struct {
		node    losslessxml.Element
		owner   string
		order   map[string]int
		updates map[string][]byte
	}{{tc, "tcPr", wordCellOrder, cellUpdates}, {tr, "trPr", wordRowOrder, rowUpdates}, {tbl, "tblPr", wordTableOrder, tableUpdates}} {
		if len(item.updates) == 0 {
			continue
		}
		raw, er := wordTableRaw(source, item.node)
		if er != nil {
			return er
		}
		updated, er := wordTableRender(raw, item.owner, item.order, item.updates)
		if er != nil {
			return er
		}
		a, b := item.node.SourceRange()
		edits = append(edits, edit{a, b, updated})
	}
	if len(edits) == 0 {
		return wordTableRefusal("invalid-table-properties", "no edit")
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].b && edits[j].a < edits[i].b {
				return wordTableRefusal("invalid-table-properties", "overlap")
			}
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	out := bytes.Clone(source)
	for _, item := range edits {
		out = append(append(bytes.Clone(out[:item.a]), item.v...), out[item.b:]...)
	}
	if _, e = losslessxml.Parse(out); e != nil {
		return wordTableRefusal("unsupported_structure", "saved XML")
	}
	if bytes.Equal(out, source) {
		return nil
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: t.part, ExpectedSHA256: t.hash, Data: out}}); e != nil {
		return e
	}
	s.generation++
	return nil
}

// Table editing refuses protection and review markup before acquiring a target.
func (s *EditSession) wordTableProtection(d *losslessxml.Document) error {
	g, e := s.pkg.Graph()
	if e != nil {
		return e
	}
	settingsCount := 0
	for _, edge := range g.Edges {
		if edge.Source != s.part || edge.Type != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" {
			continue
		}
		settingsCount++
		if settingsCount != 1 || edge.External || edge.ResolvedPart == "" {
			return wordTableRefusal("unsupported_structure", "ambiguous/external settings")
		}
		settingsType := ""
		for _, part := range g.Parts {
			if part.Name == edge.ResolvedPart {
				settingsType = part.ContentType
				break
			}
		}
		if settingsType != packaging.ContentTypeSettings {
			return wordTableRefusal("unsupported_structure", "settings content type")
		}
		blob, _, er := s.pkg.Part(edge.ResolvedPart)
		if er != nil {
			return er
		}
		settings, er := losslessxml.Parse(blob)
		if er != nil {
			return er
		}
		if len(settings.Elements()) == 0 || settings.Elements()[0].Name() != retainedName("settings") {
			return wordTableRefusal("unsupported_structure", "settings root")
		}
		for _, n := range settings.Elements() {
			if n.Name() == retainedName("trackRevisions") {
				return wordTableRefusal("unsupported_structure", "tracked revisions")
			}
			if n.Name() == retainedName("documentProtection") {
				switch retainedAttr(n, "enforcement") {
				case "0", "false", "off":
				case "1", "true", "on":
					return wordTableRefusal("protected_operation", "document protection")
				default:
					return wordTableRefusal("unsupported_structure", "unknown protection enforcement")
				}
			}
		}
	}
	blocked := map[string]bool{"ins": true, "del": true, "moveFrom": true, "moveTo": true, "fldChar": true, "fldSimple": true, "instrText": true, "commentRangeStart": true, "commentRangeEnd": true, "bookmarkStart": true, "bookmarkEnd": true, "permStart": true, "permEnd": true}
	for _, n := range d.Elements() {
		if n.Name().Space == retainedW && (blocked[n.Name().Local] || len(n.Name().Local) >= 6 && n.Name().Local[len(n.Name().Local)-6:] == "Change") {
			return wordTableRefusal("unsupported_structure", "review/field/range markup")
		}
	}
	return nil
}

// wordTableAdmit checks the entire selected physical grid and owner topology.
func wordTableAdmit(d *losslessxml.Document, tbl losslessxml.Element) error {
	if e := wordTableDescendants(d, tbl); e != nil {
		return e
	}
	if len(tbl.Attributes()) != 0 {
		return wordTableRefusal("unsupported_structure", "table attributes")
	}
	grid, e := wordTableOne(d, tbl, "tblGrid")
	if e != nil {
		return e
	}
	if len(grid.Attributes()) != 0 {
		return wordTableRefusal("unsupported_structure", "grid attributes")
	}
	for _, n := range d.Elements() {
		p, ok := n.Parent()
		if !ok {
			continue
		}
		if p == tbl && n.Name() != retainedName("tblPr") && n.Name() != retainedName("tblGrid") && n.Name() != retainedName("tr") {
			return wordTableRefusal("unsupported_structure", "table direct child")
		}
		if p == grid && n.Name() != retainedName("gridCol") {
			return wordTableRefusal("unsupported_structure", "grid direct child")
		}
	}
	cols := wordTableChildren(d, grid, "gridCol")
	if len(cols) == 0 {
		return wordTableRefusal("unsupported_structure", "empty grid")
	}
	for _, col := range cols {
		if e = wordTableLeaf(d, col, "w"); e != nil {
			return e
		}
		if e = wordTableScalar(col, "w", 1, 31680); e != nil {
			return e
		}
	}
	rows := wordTableChildren(d, tbl, "tr")
	if len(rows) == 0 {
		return wordTableRefusal("unsupported_structure", "empty table")
	}
	for _, row := range rows {
		for _, n := range d.Elements() {
			p, ok := n.Parent()
			if ok && p == row && n.Name() != retainedName("trPr") && n.Name() != retainedName("tc") {
				return wordTableRefusal("unsupported_structure", "row direct child")
			}
		}
		for _, attr := range row.Attributes() {
			if !(attr.Name.Space == retainedW && attr.Name.Local == "rsidR") && !(attr.Name.Space == "http://schemas.microsoft.com/office/word/2010/wordml" && (attr.Name.Local == "paraId" || attr.Name.Local == "textId")) {
				return wordTableRefusal("unsupported_structure", "row attributes")
			}
		}
		cells := wordTableChildren(d, row, "tc")
		if len(cells) != len(cols) {
			return wordTableRefusal("unsupported_structure", "nonrectangular table")
		}
		for _, cell := range cells {
			if len(cell.Attributes()) != 0 {
				return wordTableRefusal("unsupported_structure", "merged/foreign cell")
			}
			if len(wordTableChildren(d, cell, "tbl")) > 0 {
				return wordTableRefusal("unsupported_structure", "nested table")
			}
			if len(wordTableChildren(d, cell, "tcPr")) != 1 {
				return wordTableRefusal("unsupported_structure", "cell properties")
			}
			for _, n := range d.Elements() {
				p, ok := n.Parent()
				if ok && p == cell && n.Name() != retainedName("tcPr") && n.Name() != retainedName("p") {
					return wordTableRefusal("unsupported_structure", "cell child profile")
				}
			}
			props := wordTableChildren(d, cell, "tcPr")[0]
			children, e := wordTableAllowed(d, props, wordCellOrder)
			if e != nil {
				return e
			}
			if w, ok := children["tcW"]; ok {
				if e = wordTableLeaf(d, w, "w", "type"); e != nil {
					return e
				}
				if e = wordTableEnum(w, "type", "dxa"); e != nil {
					return e
				}
				if e = wordTableScalar(w, "w", 0, 31680); e != nil {
					return e
				}
			}
			if _, ok := children["gridSpan"]; ok {
				return wordTableRefusal("unsupported_structure", "grid span")
			}
			if _, ok := children["vMerge"]; ok {
				return wordTableRefusal("unsupported_structure", "vertical merge")
			}
		}
	}
	return nil
}

func wordTableDescendants(d *losslessxml.Document, tbl losslessxml.Element) error {
	a, b := tbl.SourceRange()
	blocked := map[string]bool{"hyperlink": true, "sdt": true, "fldChar": true, "fldSimple": true, "instrText": true, "ins": true, "del": true, "moveFrom": true, "moveTo": true, "commentRangeStart": true, "commentRangeEnd": true, "bookmarkStart": true, "bookmarkEnd": true, "permStart": true, "permEnd": true}
	for _, n := range d.Elements() {
		start, end := n.SourceRange()
		if start < a || end > b {
			continue
		}
		name := n.Name()
		if name.Space != retainedW || n != tbl && name.Local == "tbl" || blocked[name.Local] || strings.HasSuffix(name.Local, "Change") {
			return wordTableRefusal("unsupported_structure", "foreign/field/review table descendant")
		}
	}
	return nil
}
