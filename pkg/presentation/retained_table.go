package presentation

import (
	"bytes"
	"encoding/xml"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// RetainedTableTarget is a single-owner snapshot of one ordinary table frame.
type RetainedTableTarget struct {
	session            *EditSession
	generation         uint64
	part, hash         string
	frameID            uint32
	doc                *losslessxml.Document
	frame, table, cell losslessxml.Element
}

func tableRefuse(kind, why string) error { return editRefusal(kind, why) }
func tableChildren(d *losslessxml.Document, p losslessxml.Element, ns, local string) []losslessxml.Element {
	return fmtChildren(d, p, xml.Name{Space: ns, Local: local})
}
func tableOne(d *losslessxml.Document, p losslessxml.Element, ns, local string) (losslessxml.Element, error) {
	return fmtOne(d, p, xml.Name{Space: ns, Local: local})
}
func tableAttr(e losslessxml.Element, k string) (string, bool) {
	for _, a := range e.Attributes() {
		if a.Name.Space == "" && a.Name.Local == k {
			return a.Value, true
		}
	}
	return "", false
}
func tableBound(v any, min, max int64) (int64, error) {
	var n int64
	switch x := v.(type) {
	case int:
		n = int64(x)
	case int64:
		n = x
	case float64:
		if x != float64(int64(x)) {
			return 0, tableRefuse("invalid-table-properties", "integer required")
		}
		n = int64(x)
	default:
		return 0, tableRefuse("invalid-table-properties", "integer required")
	}
	if n < min || n > max {
		return 0, tableRefuse("invalid-table-properties", "integer bounds")
	}
	return n, nil
}
func tableString(v any, choices ...string) (string, error) {
	s, ok := v.(string)
	if !ok {
		return "", tableRefuse("invalid-table-properties", "string required")
	}
	for _, x := range choices {
		if x == s {
			return s, nil
		}
	}
	return "", tableRefuse("invalid-table-properties", "unsupported choice")
}
func tableBool(v any) (string, error) {
	b, ok := v.(bool)
	if !ok {
		return "", tableRefuse("invalid-table-properties", "boolean required")
	}
	if b {
		return "1", nil
	}
	return "0", nil
}
func tableNumber(e losslessxml.Element, k string, min, max int64) (int64, error) {
	s, ok := tableAttr(e, k)
	if !ok {
		return 0, tableRefuse("unsupported_structure", "missing "+k)
	}
	n, err := strconv.ParseInt(s, 10, 64)
	if err != nil || n < min || n > max {
		return 0, tableRefuse("unsupported_structure", "invalid original "+k)
	}
	return n, nil
}
func tableSafeAttrs(e losslessxml.Element, allow ...string) error {
	allowed := map[string]bool{}
	for _, v := range allow {
		allowed[v] = true
	}
	seen := map[string]bool{}
	for _, a := range e.Attributes() {
		if a.Name.Space != "" || !allowed[a.Name.Local] || seen[a.Name.Local] {
			return tableRefuse("unsupported_structure", "foreign or duplicate property attribute")
		}
		seen[a.Name.Local] = true
	}
	return nil
}
func tableRaw(source []byte, e losslessxml.Element) ([]byte, error) {
	a, b := e.SourceRange()
	if a < 0 || b > len(source) || a >= b {
		return nil, tableRefuse("unsupported_structure", "invalid source range")
	}
	return bytes.Clone(source[a:b]), nil
}
func tableReplace(source []byte, e losslessxml.Element, raw []byte) ([]byte, error) {
	a, b := e.SourceRange()
	if a < 0 || b > len(source) || a > b {
		return nil, tableRefuse("unsupported_structure", "invalid replacement range")
	}
	return append(append(bytes.Clone(source[:a]), raw...), source[b:]...), nil
}
func tableWrap(raw []byte) (*losslessxml.Document, losslessxml.Element, int, error) {
	prefix := []byte(`<root xmlns:a="` + packaging.NSDrawingML + `" xmlns:p="` + packaging.NSPresentationML + `">`)
	d, e := losslessxml.Parse(append(append(bytes.Clone(prefix), raw...), []byte(`</root>`)...))
	if e != nil {
		return nil, losslessxml.Element{}, 0, tableRefuse("unsupported_structure", "property parse")
	}
	if len(d.Elements()) < 2 {
		return nil, losslessxml.Element{}, 0, tableRefuse("unsupported_structure", "property root")
	}
	return d, d.Elements()[1], len(prefix), nil
}
func tableLocalPatch(raw []byte, local string, fn func(*losslessxml.Document, losslessxml.Element, []byte) ([]byte, error)) ([]byte, error) {
	d, p, _, e := tableWrap(raw)
	if e != nil {
		return nil, e
	}
	if p.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: local}) {
		return nil, tableRefuse("unsupported_structure", "wrong property root")
	}
	return fn(d, p, raw)
}
func tablePatchNode(raw []byte, node losslessxml.Element, repl []byte, offset int) ([]byte, error) {
	a, b := node.SourceRange()
	a -= offset
	b -= offset
	if a < 0 || b > len(raw) || a > b {
		return nil, tableRefuse("unsupported_structure", "invalid nested range")
	}
	return append(append(bytes.Clone(raw[:a]), repl...), raw[b:]...), nil
}

func (s *EditSession) FindRetainedTable(part string, frameID uint32, row, column int) (*RetainedTableTarget, error) {
	return s.findRetainedTable(part, frameID, row, column, true)
}

// FindContractTableCell selects a directly owned table cell without treating
// unsupported merge topology as an invalid package. The edit, not selection,
// must classify a merged or malformed table; this never weakens legacy edits.
func (s *EditSession) FindContractTableCell(part string, frameID uint32, row, column int) (*ContractTableCellTarget, error) {
	selected, err := s.findRetainedTable(part, frameID, row, column, false)
	if err != nil {
		return nil, err
	}
	issued := &ContractTableCellTarget{selected: selected}
	if s.contractTableCells == nil {
		s.contractTableCells = make(map[*ContractTableCellTarget]struct{})
	}
	s.contractTableCells[issued] = struct{}{}
	return issued, nil
}

func (s *EditSession) findRetainedTable(part string, frameID uint32, row, column int, requireOrdinaryTable bool) (*RetainedTableTarget, error) {
	if frameID == 0 || row < 0 || column < 0 {
		return nil, tableRefuse("missing_target", "table selector")
	}
	if !s.slides[part] {
		return nil, tableRefuse("missing_target", "unenrolled slide")
	}
	if e := s.manipulationProtection(); e != nil {
		return nil, e
	}
	source, hash, e := s.pkg.Part(part)
	if e != nil {
		return nil, e
	}
	d, e := losslessxml.Parse(source)
	if e != nil {
		return nil, e
	}
	nodes := d.Elements()
	if len(nodes) == 0 || nodes[0].Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sld"}) {
		return nil, tableRefuse("unsupported_structure", "slide root")
	}
	tree, e := tableOne(d, nodes[0], packaging.NSPresentationML, "cSld")
	if e != nil {
		return nil, e
	}
	tree, e = tableOne(d, tree, packaging.NSPresentationML, "spTree")
	if e != nil {
		return nil, e
	}
	var frame losslessxml.Element
	seen := map[uint32]bool{}
	for _, n := range nodes {
		if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "cNvPr"}) {
			continue
		}
		v, ok := tableAttr(n, "id")
		if !ok {
			return nil, tableRefuse("unsupported_structure", "frame ID")
		}
		id, er := strconv.ParseUint(v, 10, 32)
		if er != nil || id == 0 || seen[uint32(id)] {
			return nil, tableRefuse("ambiguous_target", "duplicate/invalid object ID")
		}
		seen[uint32(id)] = true
		if uint32(id) != frameID {
			continue
		}
		owner, ok := n.Parent()
		if !ok {
			continue
		}
		owner, ok = owner.Parent()
		if !ok {
			continue
		}
		p, ok := owner.Parent()
		if !ok || p != tree || owner.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "graphicFrame"}) {
			return nil, tableRefuse("unsupported_structure", "not a direct table frame")
		}
		frame = owner
	}
	if frame.Ordinal() < 0 {
		return nil, tableRefuse("missing_target", "frame ID")
	}
	graphic, e := tableOne(d, frame, packaging.NSDrawingML, "graphic")
	if e != nil {
		return nil, e
	}
	data, e := tableOne(d, graphic, packaging.NSDrawingML, "graphicData")
	if e != nil {
		return nil, e
	}
	if v, ok := tableAttr(data, "uri"); !ok || v != "http://schemas.openxmlformats.org/drawingml/2006/table" {
		return nil, tableRefuse("unsupported_structure", "graphic kind")
	}
	tbl, e := tableOne(d, data, packaging.NSDrawingML, "tbl")
	if e != nil {
		return nil, e
	}
	rows := tableChildren(d, tbl, packaging.NSDrawingML, "tr")
	grid, e := tableOne(d, tbl, packaging.NSDrawingML, "tblGrid")
	if e != nil {
		return nil, e
	}
	cols := tableChildren(d, grid, packaging.NSDrawingML, "gridCol")
	if len(rows) == 0 || len(cols) == 0 || row >= len(rows) || column >= len(cols) {
		return nil, tableRefuse("missing_target", "row/column")
	}
	if requireOrdinaryTable {
		if e = tableAdmit(d, source, frame, tbl, rows, cols); e != nil {
			return nil, e
		}
	}
	cells := tableChildren(d, rows[row], packaging.NSDrawingML, "tc")
	if column >= len(cells) {
		return nil, tableRefuse("missing_target", "selected table cell")
	}
	cell := cells[column]
	return &RetainedTableTarget{session: s, generation: s.generation, part: part, hash: hash, frameID: frameID, doc: d, frame: frame, table: tbl, cell: cell}, nil
}
func (s *EditSession) tableCurrent(t *RetainedTableTarget) error {
	if t == nil || t.session != s || t.generation != s.generation {
		return tableRefuse("stale_target", "foreign or stale table")
	}
	_, hash, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	if hash != t.hash {
		return tableRefuse("stale_target", "table part changed")
	}
	return nil
}

var tableLineOrder = map[string]int{"lnL": 1, "lnR": 2, "lnT": 3, "lnB": 4, "lnTlToBr": 5, "lnBlToTr": 6, "cell3D": 7, "solidFill": 8, "noFill": 8}

func tableValidateCell(d *losslessxml.Document, tc losslessxml.Element) error {
	if e := tableSafeAttrs(tc, "marL", "marR", "marT", "marB", "vert", "anchor", "anchorCtr"); e != nil {
		return e
	}
	for _, k := range []string{"marL", "marR", "marT", "marB"} {
		if _, e := tableNumber(tc, k, 0, 100000000); e != nil {
			return e
		}
	}
	for _, item := range []struct {
		key    string
		values []string
	}{{"vert", []string{"horz", "vert"}}, {"anchor", []string{"t", "ctr"}}, {"anchorCtr", []string{"0", "1"}}} {
		v, ok := tableAttr(tc, item.key)
		if !ok {
			return tableRefuse("unsupported_structure", "missing cell attribute")
		}
		found := false
		for _, x := range item.values {
			found = found || v == x
		}
		if !found {
			return tableRefuse("unsupported_structure", "invalid cell attribute")
		}
	}
	prev := 0
	seen := map[string]bool{}
	fill := 0
	for _, n := range d.Elements() {
		p, ok := n.Parent()
		if !ok || p != tc {
			continue
		}
		k := n.Name().Local
		if n.Name().Space != packaging.NSDrawingML || tableLineOrder[k] == 0 || seen[k] || tableLineOrder[k] < prev {
			return tableRefuse("unsupported_structure", "cell child order")
		}
		seen[k] = true
		prev = tableLineOrder[k]
		if k == "solidFill" || k == "noFill" {
			fill++
			if e := retainedValidateFill(d, n); e != nil {
				return e
			}
		}
	}
	if fill > 1 {
		return tableRefuse("unsupported_structure", "duplicate fill")
	}
	for _, k := range []string{"lnL", "lnR", "lnT", "lnB"} {
		nodes := tableChildren(d, tc, packaging.NSDrawingML, k)
		if len(nodes) != 1 {
			return tableRefuse("unsupported_structure", "four direct lines required")
		}
		n := nodes[0]
		if e := tableSafeAttrs(n, "w"); e != nil {
			return e
		}
		if _, e := tableNumber(n, "w", 0, 20116800); e != nil {
			return e
		}
		children, e := retainedDirect(d, n, map[string]bool{"solidFill": true, "noFill": true, "prstDash": true})
		if e != nil {
			return e
		}
		if v, ok := children["solidFill"]; ok {
			if e := retainedValidateFill(d, v); e != nil {
				return e
			}
		}
		if v, ok := children["noFill"]; ok {
			if e := retainedValidateFill(d, v); e != nil {
				return e
			}
		}
		if _, a := children["solidFill"]; a {
			if _, b := children["noFill"]; b {
				return tableRefuse("unsupported_structure", "two line fills")
			}
		}
		dash, ok := children["prstDash"]
		if !ok || len(dash.Attributes()) != 1 || dash.Attributes()[0].Name.Space != "" || dash.Attributes()[0].Name.Local != "val" || dash.Attributes()[0].Value != "solid" {
			return tableRefuse("unsupported_structure", "dash profile")
		}
	}
	return nil
}
func tablePatchCell(raw []byte, patch map[string]any) ([]byte, error) {
	return tableLocalPatch(raw, "tcPr", func(d *losslessxml.Document, tc losslessxml.Element, source []byte) ([]byte, error) {
		if e := tableValidateCell(d, tc); e != nil {
			return nil, e
		}
		if v, ok := patch["fill"]; ok {
			var fill []byte
			if v != nil {
				val, er := tableString(v, "none")
				if er != nil {
					if s, yes := v.(string); yes && retainedColor(s) {
						val = s
					} else {
						return nil, er
					}
				}
				fill, er = retainedChoice(val)
				if er != nil {
					return nil, tableRefuse("invalid-table-properties", "fill")
				}
			}
			var er error
			source, er = retainedPatchChild(source, []string{"solidFill", "noFill"}, fill, nil, map[string]bool{"lnL": true, "lnR": true, "lnT": true, "lnB": true, "solidFill": true, "noFill": true})
			if er != nil {
				return nil, er
			}
		}
		if rawBorder, ok := patch["border"]; ok {
			b, ok := rawBorder.(map[string]any)
			if !ok {
				return nil, tableRefuse("invalid-table-properties", "border")
			}
			side, er := tableString(b["side"], "left", "right", "top", "bottom")
			if er != nil {
				return nil, er
			}
			for k := range b {
				if k != "side" && k != "color" && k != "width" && k != "remove" {
					return nil, tableRefuse("invalid-table-properties", "border key")
				}
			}
			name := map[string]string{"left": "lnL", "right": "lnR", "top": "lnT", "bottom": "lnB"}[side]
			d, tc, off, er := tableWrap(source)
			if er != nil {
				return nil, er
			}
			lines := tableChildren(d, tc, packaging.NSDrawingML, name)
			if len(lines) != 1 {
				return nil, tableRefuse("unsupported_structure", "border missing")
			}
			line := lines[0]
			a, z := line.SourceRange()
			lineRaw := bytes.Clone(source[a-off : z-off])
			if remove, found := b["remove"]; found {
				if remove != true || len(b) != 2 {
					return nil, tableRefuse("invalid-table-properties", "remove border")
				}
				return tablePatchNode(source, line, nil, off)
			}
			changes := map[string]*string{}
			if value, found := b["width"]; found {
				n, er := tableBound(value, 0, 20116800)
				if er != nil {
					return nil, er
				}
				s := strconv.FormatInt(n, 10)
				changes["w"] = &s
			}
			if len(changes) > 0 {
				lineRaw, er = fmtOpening(lineRaw, changes)
				if er != nil {
					return nil, er
				}
			}
			if value, found := b["color"]; found {
				color, yes := value.(string)
				if !yes {
					return nil, tableRefuse("invalid-table-properties", "border colour")
				}
				if color != "none" && !retainedColor(color) {
					return nil, tableRefuse("invalid-table-properties", "border colour")
				}
				var fill []byte
				fill, er = retainedChoice(color)
				if er != nil {
					return nil, er
				}
				lineRaw, er = retainedPatchChild(lineRaw, []string{"solidFill", "noFill"}, fill, []string{"prstDash"}, map[string]bool{"solidFill": true, "noFill": true, "prstDash": true})
				if er != nil {
					return nil, er
				}
			}
			source, er = tablePatchNode(source, line, lineRaw, off)
			if er != nil {
				return nil, er
			}
		}
		attrs := map[string]*string{}
		if rawMargins, ok := patch["margins"]; ok {
			m, yes := rawMargins.(map[string]any)
			if !yes || len(m) != 4 {
				return nil, tableRefuse("invalid-table-properties", "margins")
			}
			for _, k := range []string{"left", "right", "top", "bottom"} {
				n, e := tableBound(m[k], 0, 100000000)
				if e != nil {
					return nil, e
				}
				v := strconv.FormatInt(n, 10)
				attrs[map[string]string{"left": "marL", "right": "marR", "top": "marT", "bottom": "marB"}[k]] = &v
			}
		}
		if x, ok := patch["anchor"]; ok {
			v, e := tableString(x, "t", "ctr")
			if e != nil {
				return nil, e
			}
			attrs["anchor"] = &v
		}
		if x, ok := patch["vert"]; ok {
			v, e := tableString(x, "horz", "vert")
			if e != nil {
				return nil, e
			}
			attrs["vert"] = &v
		}
		if x, ok := patch["anchorCtr"]; ok {
			v, e := tableBool(x)
			if e != nil {
				return nil, e
			}
			attrs["anchorCtr"] = &v
		}
		if len(attrs) > 0 {
			var er error
			source, er = fmtOpening(source, attrs)
			if er != nil {
				return nil, er
			}
		}
		return source, nil
	})
}

// SetRetainedTable edits only the selected table property spans after full preflight.
func (s *EditSession) SetRetainedTable(t *RetainedTableTarget, patch map[string]any) error {
	if e := s.tableCurrent(t); e != nil {
		return e
	}
	if len(patch) == 0 {
		return tableRefuse("invalid-table-properties", "empty patch")
	}
	valid := map[string]bool{"fill": true, "border": true, "margins": true, "anchor": true, "vert": true, "anchorCtr": true, "column": true, "width": true, "row": true, "height": true, "firstRow": true, "bandRow": true, "x": true, "y": true, "bold": true}
	for k := range patch {
		if !valid[k] {
			return tableRefuse("invalid-table-properties", "unknown patch key")
		}
	}
	source, _, e := s.pkg.Part(t.part)
	if e != nil {
		return e
	}
	d := t.doc
	frame := t.frame
	table := t.table
	cell := t.cell
	tc, e := tableOne(d, cell, packaging.NSDrawingML, "tcPr")
	if e != nil {
		return e
	}
	if e = tableValidateCell(d, tc); e != nil {
		return e
	}
	grid, e := tableOne(d, table, packaging.NSDrawingML, "tblGrid")
	if e != nil {
		return e
	}
	if e = tableAdmit(d, source, frame, table, tableChildren(d, table, packaging.NSDrawingML, "tr"), tableChildren(d, grid, packaging.NSDrawingML, "gridCol")); e != nil {
		return e
	}
	// A planned edit changes disjoint original source ranges only.
	type change struct {
		a, b int
		data []byte
	}
	edits := []change{}
	add := func(n losslessxml.Element, v []byte) { a, b := n.SourceRange(); edits = append(edits, change{a, b, v}) }
	cellPatch := false
	for _, key := range []string{"fill", "border", "margins", "anchor", "vert", "anchorCtr"} {
		if _, ok := patch[key]; ok {
			cellPatch = true
		}
	}
	if cellPatch {
		chosen := map[string]any{}
		for _, key := range []string{"fill", "border", "margins", "anchor", "vert", "anchorCtr"} {
			if v, ok := patch[key]; ok {
				chosen[key] = v
			}
		}
		raw, er := tableRaw(source, tc)
		if er != nil {
			return er
		}
		out, er := tablePatchCell(raw, chosen)
		if er != nil {
			return er
		}
		add(tc, out)
	}
	if _, ok := patch["column"]; ok || patch["width"] != nil {
		col, er := tableBound(patch["column"], 0, 1000000)
		if er != nil {
			return er
		}
		width, er := tableBound(patch["width"], 1, 100000000)
		if er != nil {
			return er
		}
		grid, er := tableOne(d, table, packaging.NSDrawingML, "tblGrid")
		if er != nil {
			return er
		}
		cols := tableChildren(d, grid, packaging.NSDrawingML, "gridCol")
		if int(col) >= len(cols) {
			return tableRefuse("missing_target", "grid column")
		}
		var sum int64
		for _, x := range cols {
			if er = tableSafeAttrs(x, "w"); er != nil {
				return er
			}
			n, ex := tableNumber(x, "w", 1, 100000000)
			if ex != nil {
				return ex
			}
			sum += n
		}
		xfrm, er := tableOne(d, frame, packaging.NSPresentationML, "xfrm")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(xfrm); er != nil {
			return er
		}
		ext, er := tableOne(d, xfrm, packaging.NSDrawingML, "ext")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(ext, "cx", "cy"); er != nil {
			return er
		}
		cx, er := tableNumber(ext, "cx", 1, 100000000)
		if er != nil {
			return er
		}
		if sum != cx {
			return tableRefuse("unsupported_structure", "frame grid extent mismatch")
		}
		old, er := tableNumber(cols[col], "w", 1, 100000000)
		if er != nil {
			return er
		}
		total := sum - old + width
		off, er := tableOne(d, xfrm, packaging.NSDrawingML, "off")
		if er != nil {
			return er
		}
		x, er := tableNumber(off, "x", 0, 100000000)
		if er != nil {
			return er
		}
		if total < 1 || total > 100000000 || x+total > 100000000 {
			return tableRefuse("invalid-table-properties", "frame width")
		}
		raw, _ := tableRaw(source, cols[col])
		v := strconv.FormatInt(width, 10)
		raw, er = fmtOpening(raw, map[string]*string{"w": &v})
		if er != nil {
			return er
		}
		add(cols[col], raw)
		raw, _ = tableRaw(source, ext)
		v = strconv.FormatInt(total, 10)
		raw, er = fmtOpening(raw, map[string]*string{"cx": &v})
		if er != nil {
			return er
		}
		add(ext, raw)
	}
	if _, ok := patch["row"]; ok || patch["height"] != nil {
		row, er := tableBound(patch["row"], 0, 1000000)
		if er != nil {
			return er
		}
		height, er := tableBound(patch["height"], 1, 100000000)
		if er != nil {
			return er
		}
		rows := tableChildren(d, table, packaging.NSDrawingML, "tr")
		if int(row) >= len(rows) {
			return tableRefuse("missing_target", "table row")
		}
		var sum int64
		for _, x := range rows {
			if er = tableSafeAttrs(x, "h"); er != nil {
				return er
			}
			n, ex := tableNumber(x, "h", 1, 100000000)
			if ex != nil {
				return ex
			}
			sum += n
		}
		xfrm, er := tableOne(d, frame, packaging.NSPresentationML, "xfrm")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(xfrm); er != nil {
			return er
		}
		ext, er := tableOne(d, xfrm, packaging.NSDrawingML, "ext")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(ext, "cx", "cy"); er != nil {
			return er
		}
		cy, er := tableNumber(ext, "cy", 1, 100000000)
		if er != nil {
			return er
		}
		if sum != cy {
			return tableRefuse("unsupported_structure", "frame row extent mismatch")
		}
		old, er := tableNumber(rows[row], "h", 1, 100000000)
		if er != nil {
			return er
		}
		total := sum - old + height
		off, er := tableOne(d, xfrm, packaging.NSDrawingML, "off")
		if er != nil {
			return er
		}
		y, er := tableNumber(off, "y", 0, 100000000)
		if er != nil {
			return er
		}
		if total < 1 || total > 100000000 || y+total > 100000000 {
			return tableRefuse("invalid-table-properties", "frame height")
		}
		raw, _ := tableRaw(source, rows[row])
		v := strconv.FormatInt(height, 10)
		raw, er = fmtOpening(raw, map[string]*string{"h": &v})
		if er != nil {
			return er
		}
		add(rows[row], raw)
		raw, _ = tableRaw(source, ext)
		v = strconv.FormatInt(total, 10)
		raw, er = fmtOpening(raw, map[string]*string{"cy": &v})
		if er != nil {
			return er
		}
		add(ext, raw)
	}
	for _, key := range []string{"firstRow", "bandRow"} {
		if x, ok := patch[key]; ok {
			v, er := tableBool(x)
			if er != nil {
				return er
			}
			pr, er := tableOne(d, table, packaging.NSDrawingML, "tblPr")
			if er != nil {
				return er
			}
			if er = tableSafeAttrs(pr, "firstRow", "bandRow"); er != nil {
				return er
			}
			for _, k := range []string{"firstRow", "bandRow"} {
				s, yes := tableAttr(pr, k)
				if !yes || s != "0" && s != "1" {
					return tableRefuse("unsupported_structure", "invalid original table flag")
				}
			}
			raw, _ := tableRaw(source, pr)
			raw, er = fmtOpening(raw, map[string]*string{key: &v})
			if er != nil {
				return er
			}
			add(pr, raw)
		}
	}
	if _, ok := patch["x"]; ok || patch["y"] != nil {
		x, er := tableBound(patch["x"], 0, 100000000)
		if er != nil {
			return er
		}
		y, er := tableBound(patch["y"], 0, 100000000)
		if er != nil {
			return er
		}
		xf, er := tableOne(d, frame, packaging.NSPresentationML, "xfrm")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(xf); er != nil {
			return er
		}
		off, er := tableOne(d, xf, packaging.NSDrawingML, "off")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(off, "x", "y"); er != nil {
			return er
		}
		ext, er := tableOne(d, xf, packaging.NSDrawingML, "ext")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(ext, "cx", "cy"); er != nil {
			return er
		}
		cx, er := tableNumber(ext, "cx", 1, 100000000)
		if er != nil {
			return er
		}
		cy, er := tableNumber(ext, "cy", 1, 100000000)
		if er != nil {
			return er
		}
		if _, er = tableNumber(off, "x", 0, 100000000); er != nil {
			return er
		}
		if _, er = tableNumber(off, "y", 0, 100000000); er != nil {
			return er
		}
		if x+cx > 100000000 || y+cy > 100000000 {
			return tableRefuse("invalid-table-properties", "frame position")
		}
		raw, _ := tableRaw(source, off)
		sx, sy := strconv.FormatInt(x, 10), strconv.FormatInt(y, 10)
		raw, er = fmtOpening(raw, map[string]*string{"x": &sx, "y": &sy})
		if er != nil {
			return er
		}
		add(off, raw)
	}
	if value, ok := patch["bold"]; ok {
		v, er := tableBool(value)
		if er != nil {
			return er
		}
		body, er := tableOne(d, cell, packaging.NSDrawingML, "txBody")
		if er != nil {
			return er
		}
		ps := tableChildren(d, body, packaging.NSDrawingML, "p")
		if len(ps) == 0 {
			return tableRefuse("unsupported_structure", "cell paragraph")
		}
		runs := tableChildren(d, ps[0], packaging.NSDrawingML, "r")
		if len(runs) == 0 {
			return tableRefuse("unsupported_structure", "cell run")
		}
		rp, er := tableOne(d, runs[0], packaging.NSDrawingML, "rPr")
		if er != nil {
			return er
		}
		if er = tableSafeAttrs(rp, "b"); er != nil {
			return er
		}
		orig, yes := tableAttr(rp, "b")
		if !yes || orig != "0" && orig != "1" {
			return tableRefuse("unsupported_structure", "original bold")
		}
		raw, _ := tableRaw(source, rp)
		raw, er = fmtOpening(raw, map[string]*string{"b": &v})
		if er != nil {
			return er
		}
		add(rp, raw)
	}
	if len(edits) == 0 {
		return tableRefuse("invalid-table-properties", "no property")
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].b && edits[j].a < edits[i].b {
				return tableRefuse("invalid-table-properties", "overlapping table patches")
			}
		}
	}
	for i := 0; i < len(edits); i++ {
		for j := i + 1; j < len(edits); j++ {
			if edits[i].a < edits[j].a {
				edits[i], edits[j] = edits[j], edits[i]
			}
		}
	}
	out := bytes.Clone(source)
	for _, edit := range edits {
		out = append(append(bytes.Clone(out[:edit.a]), edit.data...), out[edit.b:]...)
	}
	if _, e = losslessxml.Parse(out); e != nil {
		return tableRefuse("unsupported_structure", "invalid saved XML")
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

// tableDirectProfile refuses unknown, duplicate and reordered direct owners.
func tableDirectProfile(d *losslessxml.Document, owner losslessxml.Element, required []xml.Name, repeated xml.Name) error {
	index, count := 0, 0
	for _, n := range d.Elements() {
		p, ok := n.Parent()
		if !ok || p != owner {
			continue
		}
		if index < len(required) && n.Name() == required[index] {
			index++
			continue
		}
		if index != len(required) || n.Name() != repeated {
			return tableRefuse("unsupported_structure", "unknown/reordered direct table owner")
		}
		count++
	}
	if index != len(required) || repeated.Local != "" && count == 0 {
		return tableRefuse("unsupported_structure", "missing direct table owner")
	}
	return nil
}

// tableAdmit checks complete frame/table ownership before any property edit.
func tableAdmit(d *losslessxml.Document, source []byte, frame, tbl losslessxml.Element, rows, cols []losslessxml.Element) error {
	if len(rows) == 0 || len(cols) == 0 {
		return tableRefuse("unsupported_structure", "empty table")
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	profiles := []struct {
		owner    losslessxml.Element
		required []xml.Name
		repeated xml.Name
	}{
		{frame, []xml.Name{p("nvGraphicFramePr"), p("xfrm"), a("graphic")}, xml.Name{}},
		{tbl, []xml.Name{a("tblPr"), a("tblGrid")}, a("tr")},
	}
	for _, profile := range profiles {
		if e := tableDirectProfile(d, profile.owner, profile.required, profile.repeated); e != nil {
			return e
		}
	}
	nv, e := tableOne(d, frame, packaging.NSPresentationML, "nvGraphicFramePr")
	if e != nil {
		return e
	}
	if e = tableDirectProfile(d, nv, []xml.Name{p("cNvPr"), p("cNvGraphicFramePr"), p("nvPr")}, xml.Name{}); e != nil {
		return e
	}
	cNv, e := tableOne(d, nv, packaging.NSPresentationML, "cNvGraphicFramePr")
	if e != nil {
		return e
	}
	locks, e := tableOne(d, cNv, packaging.NSDrawingML, "graphicFrameLocks")
	if e != nil {
		return e
	}
	if e = tableDirectProfile(d, cNv, []xml.Name{a("graphicFrameLocks")}, xml.Name{}); e != nil {
		return e
	}
	if e = tableDirectProfile(d, locks, nil, xml.Name{}); e != nil {
		return e
	}
	if e = tableSafeAttrs(locks, "noGrp"); e != nil {
		return e
	}
	if v, ok := tableAttr(locks, "noGrp"); !ok || v != "1" {
		return tableRefuse("protected_operation", "graphic frame lock profile")
	}
	xf, e := tableOne(d, frame, packaging.NSPresentationML, "xfrm")
	if e != nil {
		return e
	}
	if e = tableDirectProfile(d, xf, []xml.Name{a("off"), a("ext")}, xml.Name{}); e != nil {
		return e
	}
	if e = tableSafeAttrs(xf); e != nil {
		return e
	}
	off, e := tableOne(d, xf, packaging.NSDrawingML, "off")
	if e != nil {
		return e
	}
	ext, e := tableOne(d, xf, packaging.NSDrawingML, "ext")
	if e != nil {
		return e
	}
	if e = tableSafeAttrs(off, "x", "y"); e != nil {
		return e
	}
	if e = tableSafeAttrs(ext, "cx", "cy"); e != nil {
		return e
	}
	x, e := tableNumber(off, "x", 0, 100000000)
	if e != nil {
		return e
	}
	y, e := tableNumber(off, "y", 0, 100000000)
	if e != nil {
		return e
	}
	cx, e := tableNumber(ext, "cx", 1, 100000000)
	if e != nil {
		return e
	}
	cy, e := tableNumber(ext, "cy", 1, 100000000)
	if e != nil {
		return e
	}
	if x+cx > 100000000 || y+cy > 100000000 {
		return tableRefuse("unsupported_structure", "frame bounds")
	}
	graphic, e := tableOne(d, frame, packaging.NSDrawingML, "graphic")
	if e != nil {
		return e
	}
	if e = tableDirectProfile(d, graphic, []xml.Name{a("graphicData")}, xml.Name{}); e != nil {
		return e
	}
	data, e := tableOne(d, graphic, packaging.NSDrawingML, "graphicData")
	if e != nil {
		return e
	}
	if e = tableDirectProfile(d, data, []xml.Name{a("tbl")}, xml.Name{}); e != nil {
		return e
	}
	grid, e := tableOne(d, tbl, packaging.NSDrawingML, "tblGrid")
	if e != nil {
		return e
	}
	if e = tableDirectProfile(d, grid, nil, a("gridCol")); e != nil {
		return e
	}
	for _, row := range rows {
		if e = tableDirectProfile(d, row, nil, a("tc")); e != nil {
			return e
		}
		for _, cell := range tableChildren(d, row, packaging.NSDrawingML, "tc") {
			if e = tableDirectProfile(d, cell, []xml.Name{a("txBody"), a("tcPr")}, xml.Name{}); e != nil {
				return e
			}
		}
	}
	pr, e := tableOne(d, tbl, packaging.NSDrawingML, "tblPr")
	if e != nil {
		return e
	}
	if e = tableSafeAttrs(pr, "firstRow", "bandRow"); e != nil {
		return e
	}
	for _, key := range []string{"firstRow", "bandRow"} {
		v, ok := tableAttr(pr, key)
		if !ok || v != "0" && v != "1" {
			return tableRefuse("unsupported_structure", "original table flags")
		}
	}
	style := tableChildren(d, pr, packaging.NSDrawingML, "tableStyleId")
	if len(style) != 1 || len(style[0].Attributes()) != 0 {
		return tableRefuse("unsupported_structure", "table style child")
	}
	for _, n := range d.Elements() {
		p, ok := n.Parent()
		if ok && p == pr && n != style[0] {
			return tableRefuse("unsupported_structure", "table property child")
		}
	}
	var width, height int64
	for _, col := range cols {
		if e = tableSafeAttrs(col, "w"); e != nil {
			return e
		}
		v, er := tableNumber(col, "w", 1, 100000000)
		if er != nil {
			return er
		}
		width += v
	}
	for _, row := range rows {
		if e = tableSafeAttrs(row, "h"); e != nil {
			return e
		}
		v, er := tableNumber(row, "h", 1, 100000000)
		if er != nil {
			return er
		}
		height += v
		cells := tableChildren(d, row, packaging.NSDrawingML, "tc")
		if len(cells) != len(cols) {
			return tableRefuse("unsupported_structure", "nonrectangular table")
		}
		for _, cell := range cells {
			if len(cell.Attributes()) != 0 {
				return tableRefuse("unsupported_structure", "merged/foreign cell")
			}
			tc, er := tableOne(d, cell, packaging.NSDrawingML, "tcPr")
			if er != nil {
				return er
			}
			if len(tc.Attributes()) > 0 || len(tableChildren(d, tc, packaging.NSDrawingML, "lnL")) > 0 {
				if er = tableValidateCell(d, tc); er != nil {
					return er
				}
			} else {
				fills := 0
				for _, n := range d.Elements() {
					p, ok := n.Parent()
					if !ok || p != tc {
						continue
					}
					if n.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: "solidFill"}) && n.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: "noFill"}) {
						return tableRefuse("unsupported_structure", "unprofiled neighbour cell")
					}
					fills++
					if e := retainedValidateFill(d, n); e != nil {
						return e
					}
				}
				if fills > 1 {
					return tableRefuse("unsupported_structure", "ambiguous neighbour fill")
				}
			}
			body, er := tableOne(d, cell, packaging.NSDrawingML, "txBody")
			if er != nil {
				return er
			}
			if er = tableBodyProfile(d, body); er != nil {
				return er
			}
			raw, er := tableRaw(source, body)
			if er != nil {
				return er
			}
			if bytes.Contains(raw, []byte("<!--")) || bytes.Contains(raw, []byte("<?")) {
				return tableRefuse("unsupported_structure", "table lexical barrier")
			}
		}
	}
	if width != cx || height != cy {
		return tableRefuse("unsupported_structure", "frame grid/row sum mismatch")
	}
	return nil
}

func tableBodyProfile(d *losslessxml.Document, body losslessxml.Element) error {
	allowed := map[string]bool{"txBody": true, "bodyPr": true, "lstStyle": true, "p": true, "pPr": true, "r": true, "rPr": true, "t": true, "endParaRPr": true}
	a, b := body.SourceRange()
	for _, n := range d.Elements() {
		start, end := n.SourceRange()
		if start < a || end > b {
			continue
		}
		name := n.Name()
		if name.Space != packaging.NSDrawingML || !allowed[name.Local] {
			return tableRefuse("unsupported_structure", "foreign/field/hyperlink table content")
		}
	}
	return nil
}
