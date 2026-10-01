package acceptance

import (
	"bytes"
	"encoding/json"
	"fmt"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

type tableView struct {
	doc                                                             *losslessxml.Document
	frame, table, row, cell, tc, trPr, tblPr, grid, off, ext, runPr losslessxml.Element
}

func tableOnly(d *losslessxml.Document, owner losslessxml.Element, ns, local string) (losslessxml.Element, error) {
	n := retainedChild(d, owner, ns, local)
	if len(n) != 1 {
		return losslessxml.Element{}, fmt.Errorf("table oracle %s direct count %d", local, len(n))
	}
	return n[0], nil
}
func tableViewOf(raw []byte, format string) (tableView, error) {
	d, e := losslessxml.Parse(raw)
	if e != nil {
		return tableView{}, e
	}
	v := tableView{doc: d}
	root := d.Elements()[0]
	if format == "pptx" {
		if root.Name().Space != pptxP || root.Name().Local != "sld" {
			return v, fmt.Errorf("slide root")
		}
		c, e := tableOnly(d, root, pptxP, "cSld")
		if e != nil {
			return v, e
		}
		tree, e := tableOnly(d, c, pptxP, "spTree")
		if e != nil {
			return v, e
		}
		frames := retainedChild(d, tree, pptxP, "graphicFrame")
		for _, f := range frames {
			nv, er := tableOnly(d, f, pptxP, "nvGraphicFramePr")
			if er != nil {
				return v, er
			}
			id, er := tableOnly(d, nv, pptxP, "cNvPr")
			if er != nil {
				return v, er
			}
			if retainedAttr(id, "", "id") == "3" {
				if v.frame.Ordinal() >= 0 {
					return v, fmt.Errorf("duplicate selected frame")
				}
				v.frame = f
			}
		}
		if v.frame.Ordinal() < 0 {
			return v, fmt.Errorf("missing frame 3")
		}
		xf, e := tableOnly(d, v.frame, pptxP, "xfrm")
		if e != nil {
			return v, e
		}
		v.off, e = tableOnly(d, xf, pptxA, "off")
		if e != nil {
			return v, e
		}
		v.ext, e = tableOnly(d, xf, pptxA, "ext")
		if e != nil {
			return v, e
		}
		g, e := tableOnly(d, v.frame, pptxA, "graphic")
		if e != nil {
			return v, e
		}
		data, e := tableOnly(d, g, pptxA, "graphicData")
		if e != nil {
			return v, e
		}
		v.table, e = tableOnly(d, data, pptxA, "tbl")
		if e != nil {
			return v, e
		}
		v.grid, e = tableOnly(d, v.table, pptxA, "tblGrid")
		if e != nil {
			return v, e
		}
		v.tblPr, e = tableOnly(d, v.table, pptxA, "tblPr")
		if e != nil {
			return v, e
		}
		rows := retainedChild(d, v.table, pptxA, "tr")
		if len(rows) != 6 {
			return v, fmt.Errorf("row count %d", len(rows))
		}
		v.row = rows[0]
		cells := retainedChild(d, v.row, pptxA, "tc")
		if len(cells) != 3 {
			return v, fmt.Errorf("cell count %d", len(cells))
		}
		v.cell = cells[0]
		v.tc, e = tableOnly(d, v.cell, pptxA, "tcPr")
		if e != nil {
			return v, e
		}
		body, e := tableOnly(d, v.cell, pptxA, "txBody")
		if e != nil {
			return v, e
		}
		ps := retainedChild(d, body, pptxA, "p")
		if len(ps) < 1 {
			return v, fmt.Errorf("selected paragraph")
		}
		runs := retainedChild(d, ps[0], pptxA, "r")
		if len(runs) < 1 {
			return v, fmt.Errorf("selected run")
		}
		v.runPr, e = tableOnly(d, runs[0], pptxA, "rPr")
		return v, e
	}
	if root.Name().Space != retainedDocxW || root.Name().Local != "document" {
		return v, fmt.Errorf("document root")
	}
	body, e := tableOnly(d, root, retainedDocxW, "body")
	if e != nil {
		return v, e
	}
	tables := retainedChild(d, body, retainedDocxW, "tbl")
	if len(tables) != 1 {
		return v, fmt.Errorf("selected table count %d", len(tables))
	}
	v.table = tables[0]
	v.tblPr, e = tableOnly(d, v.table, retainedDocxW, "tblPr")
	if e != nil {
		return v, e
	}
	v.grid, e = tableOnly(d, v.table, retainedDocxW, "tblGrid")
	if e != nil {
		return v, e
	}
	rows := retainedChild(d, v.table, retainedDocxW, "tr")
	if len(rows) != 4 {
		return v, fmt.Errorf("row count %d", len(rows))
	}
	v.row = rows[0]
	v.trPr, e = tableOnly(d, v.row, retainedDocxW, "trPr")
	if e != nil {
		return v, e
	}
	cells := retainedChild(d, v.row, retainedDocxW, "tc")
	if len(cells) != 3 {
		return v, fmt.Errorf("cell count %d", len(cells))
	}
	v.cell = cells[0]
	v.tc, e = tableOnly(d, v.cell, retainedDocxW, "tcPr")
	return v, e
}
func tableVal(v any) string {
	switch x := v.(type) {
	case bool:
		if x {
			return "1"
		}
		return "0"
	case float64:
		return strconv.FormatInt(int64(x), 10)
	case string:
		return x
	}
	return ""
}
func tableExpectAttr(n losslessxml.Element, ns, key string, want any) error {
	got := retainedAttr(n, ns, key)
	if got != tableVal(want) {
		return fmt.Errorf("saved %s=%q want %q", key, got, tableVal(want))
	}
	return nil
}
func tableExpected(s *tableState) error {
	if s.saved == nil {
		return fmt.Errorf("unsaved table")
	}
	var expected map[string]any
	if e := json.Unmarshal(s.record.Expected, &expected); e != nil {
		return e
	}
	if s.record.Kind == "refusal" {
		if s.refusal == nil {
			return fmt.Errorf("missing production refusal")
		}
		return nil
	}
	v, e := tableViewOf(s.after[s.part()], s.record.Format)
	if e != nil {
		return e
	}
	one := func(p losslessxml.Element, ns, name string) (losslessxml.Element, error) {
		return tableOnly(v.doc, p, ns, name)
	}
	ns := pptxA
	if s.record.Format == "docx" {
		ns = retainedDocxW
	}
	if s.record.Format == "pptx" {
		switch s.record.Kind {
		case "cell-fill", "cell-no-fill", "cell-inherit-fill":
			if s.record.Kind == "cell-inherit-fill" {
				if len(retainedChild(v.doc, v.tc, ns, "solidFill"))+len(retainedChild(v.doc, v.tc, ns, "noFill")) != 0 {
					return fmt.Errorf("fill not inherited")
				}
				return nil
			}
			name := "solidFill"
			if s.record.Kind == "cell-no-fill" {
				name = "noFill"
			}
			fill, er := one(v.tc, ns, name)
			if er != nil {
				return er
			}
			if name == "noFill" {
				if len(retainedChild(v.doc, v.tc, ns, "solidFill")) != 0 {
					return fmt.Errorf("solid fill remains")
				}
				return nil
			}
			clr, er := one(fill, ns, "srgbClr")
			if er != nil {
				return er
			}
			return tableExpectAttr(clr, "", "val", expected["fill"])
		case "border-left", "border-right", "border-top", "border-bottom", "border-no-fill", "border-remove":
			var b map[string]any
			if x, ok := expected["border"]; ok {
				b = x.(map[string]any)
			}
			if s.record.Kind == "border-remove" {
				if len(retainedChild(v.doc, v.tc, ns, "lnL")) != 0 {
					return fmt.Errorf("left border remains")
				}
				return nil
			}
			side := b["side"].(string)
			name := map[string]string{"left": "lnL", "right": "lnR", "top": "lnT", "bottom": "lnB"}[side]
			line, er := one(v.tc, ns, name)
			if er != nil {
				return er
			}
			if er = tableExpectAttr(line, "", "w", b["width"]); er != nil {
				return er
			}
			dash, er := one(line, ns, "prstDash")
			if er != nil {
				return er
			}
			if er = tableExpectAttr(dash, "", "val", b["dash"]); er != nil {
				return er
			}
			if b["fill"] == "none" {
				_, er = one(line, ns, "noFill")
				if er != nil {
					return er
				}
				if len(retainedChild(v.doc, line, ns, "solidFill")) > 0 {
					return fmt.Errorf("border fill remains")
				}
				return nil
			}
			fill, er := one(line, ns, "solidFill")
			if er != nil {
				return er
			}
			clr, er := one(fill, ns, "srgbClr")
			if er != nil {
				return er
			}
			return tableExpectAttr(clr, "", "val", b["color"])
		case "cell-margins", "cell-anchor", "cell-direction", "cell-anchor-center":
			if s.record.Kind == "cell-margins" {
				m := expected["marL"]
				for _, key := range []string{"marL", "marR", "marT", "marB"} {
					if e = tableExpectAttr(v.tc, "", key, expected[key]); e != nil {
						return e
					}
				}
				_ = m
				return nil
			}
			key := map[string]string{"cell-anchor": "anchor", "cell-direction": "vert", "cell-anchor-center": "anchorCtr"}[s.record.Kind]
			return tableExpectAttr(v.tc, "", key, expected[key])
		case "column-width":
			cols := retainedChild(v.doc, v.grid, ns, "gridCol")
			widths := expected["columnWidths"].([]any)
			if len(cols) != len(widths) {
				return fmt.Errorf("column count")
			}
			for i, n := range cols {
				if e = tableExpectAttr(n, "", "w", widths[i]); e != nil {
					return e
				}
			}
			return tableExpectAttr(v.ext, "", "cx", expected["frameWidth"])
		case "row-height":
			rows := retainedChild(v.doc, v.table, ns, "tr")
			heights := expected["rowHeights"].([]any)
			if len(rows) != len(heights) {
				return fmt.Errorf("row count")
			}
			for i, n := range rows {
				if e = tableExpectAttr(n, "", "h", heights[i]); e != nil {
					return e
				}
			}
			return tableExpectAttr(v.ext, "", "cy", expected["frameHeight"])
		case "first-row-off", "band-row-off":
			key := map[string]string{"first-row-off": "firstRow", "band-row-off": "bandRow"}[s.record.Kind]
			return tableExpectAttr(v.tblPr, "", key, expected[key])
		case "frame-position":
			for _, key := range []string{"x", "y"} {
				if e = tableExpectAttr(v.off, "", key, expected[key]); e != nil {
					return e
				}
			}
			if e = tableExpectAttr(v.ext, "", "cx", expected["width"]); e != nil {
				return e
			}
			return tableExpectAttr(v.ext, "", "cy", expected["height"])
		case "cell-unbold":
			return tableExpectAttr(v.runPr, "", "b", expected["bold"])
		}
		return fmt.Errorf("unknown PPTX table kind %s", s.record.Kind)
	}
	switch s.record.Kind {
	case "cell-shading":
		shd, er := one(v.tc, ns, "shd")
		if er != nil {
			return er
		}
		for _, pair := range []struct{ attr, key string }{{"fill", "shading"}, {"val", "pattern"}, {"color", "color"}} {
			if er = tableExpectAttr(shd, ns, pair.attr, expected[pair.key]); er != nil {
				return er
			}
		}
		return nil
	case "cell-inherit-shading":
		if len(retainedChild(v.doc, v.tc, ns, "shd")) != 0 {
			return fmt.Errorf("shading remains")
		}
		return nil
	case "cell-anchor", "cell-direction", "cell-no-wrap", "cell-fit-text", "cell-hide-mark":
		var key, tag string
		switch s.record.Kind {
		case "cell-anchor":
			key, tag = "verticalAlign", "vAlign"
		case "cell-direction":
			key, tag = "textDirection", "textDirection"
		case "cell-no-wrap":
			key, tag = "noWrap", "noWrap"
		case "cell-fit-text":
			key, tag = "tcFitText", "tcFitText"
		case "cell-hide-mark":
			key, tag = "hideMark", "hideMark"
		}
		n, er := one(v.tc, ns, tag)
		if er != nil {
			return er
		}
		return tableExpectAttr(n, ns, "val", expected[key])
	case "cell-margins":
		parent, er := one(v.tc, ns, "tcMar")
		if er != nil {
			return er
		}
		m := expected["margins"].(map[string]any)
		for _, side := range []string{"top", "left", "bottom", "right"} {
			n, er := one(parent, ns, side)
			if er != nil {
				return er
			}
			if er = tableExpectAttr(n, ns, "w", m[side]); er != nil {
				return er
			}
			if er = tableExpectAttr(n, ns, "type", m["type"]); er != nil {
				return er
			}
		}
		return nil
	case "border-top", "border-left", "border-bottom", "border-right", "border-remove":
		parent, er := one(v.tc, ns, "tcBorders")
		if er != nil {
			return er
		}
		if s.record.Kind == "border-remove" {
			if len(retainedChild(v.doc, parent, ns, "top")) != 0 {
				return fmt.Errorf("top border remains")
			}
			return nil
		}
		b := expected["border"].(map[string]any)
		n, er := one(parent, ns, b["side"].(string))
		if er != nil {
			return er
		}
		for _, pair := range []struct{ attr, key string }{{"val", "style"}, {"sz", "size"}, {"color", "color"}} {
			if er = tableExpectAttr(n, ns, pair.attr, b[pair.key]); er != nil {
				return er
			}
		}
		return nil
	case "cell-width":
		w, er := one(v.tc, ns, "tcW")
		if er != nil {
			return er
		}
		if er = tableExpectAttr(w, ns, "w", expected["width"]); er != nil {
			return er
		}
		if er = tableExpectAttr(w, ns, "type", expected["widthType"]); er != nil {
			return er
		}
		cols := retainedChild(v.doc, v.grid, ns, "gridCol")
		grid := expected["grid"].([]any)
		if len(cols) != len(grid) {
			return fmt.Errorf("grid changed")
		}
		for i, n := range cols {
			if er = tableExpectAttr(n, ns, "w", grid[i]); er != nil {
				return er
			}
		}
		return nil
	case "row-header", "row-cant-split":
		key, tag := "tblHeader", "tblHeader"
		if s.record.Kind == "row-cant-split" {
			key, tag = "cantSplit", "cantSplit"
		}
		n, er := one(v.trPr, ns, tag)
		if er != nil {
			return er
		}
		return tableExpectAttr(n, ns, "val", expected[key])
	case "row-height":
		n, er := one(v.trPr, ns, "trHeight")
		if er != nil {
			return er
		}
		if er = tableExpectAttr(n, ns, "val", expected["height"]); er != nil {
			return er
		}
		return tableExpectAttr(n, ns, "hRule", expected["heightRule"])
	case "table-alignment":
		n, er := one(v.tblPr, ns, "jc")
		if er != nil {
			return er
		}
		return tableExpectAttr(n, ns, "val", expected["alignment"])
	case "table-indent":
		n, er := one(v.tblPr, ns, "tblInd")
		if er != nil {
			return er
		}
		if er = tableExpectAttr(n, ns, "w", expected["indent"]); er != nil {
			return er
		}
		return tableExpectAttr(n, ns, "type", expected["indentType"])
	}
	return fmt.Errorf("unknown Word table kind %s", s.record.Kind)
}

// tableOrder checks exact direct-child order separately from masked byte custody.
func tableOrder(s *tableState) error {
	if s.record.Kind == "refusal" {
		return nil
	}
	v, e := tableViewOf(s.after[s.part()], s.record.Format)
	if e != nil {
		return e
	}
	ns := pptxA
	type orderItem struct {
		n    losslessxml.Element
		rank map[string]int
	}
	items := []orderItem{}
	if s.record.Format == "pptx" {
		items = append(items, orderItem{v.tc, map[string]int{"lnL": 1, "lnR": 2, "lnT": 3, "lnB": 4, "solidFill": 5, "noFill": 5}})
		for _, side := range []string{"lnL", "lnR", "lnT", "lnB"} {
			for _, n := range retainedChild(v.doc, v.tc, ns, side) {
				items = append(items, orderItem{n, map[string]int{"solidFill": 1, "noFill": 1, "prstDash": 2}})
			}
		}
		items = append(items, orderItem{v.tblPr, map[string]int{"tableStyleId": 1}})
	} else {
		ns = retainedDocxW
		items = append(items,
			orderItem{v.tc, map[string]int{"tcW": 1, "tcBorders": 2, "shd": 3, "noWrap": 4, "tcMar": 5, "textDirection": 6, "tcFitText": 7, "vAlign": 8, "hideMark": 9}},
			orderItem{v.trPr, map[string]int{"cantSplit": 1, "trHeight": 2, "tblHeader": 3}},
			orderItem{v.tblPr, map[string]int{"tblStyle": 1, "tblW": 2, "jc": 3, "tblInd": 4, "tblLook": 5}})
		for _, name := range []string{"tcBorders", "tcMar"} {
			n, er := tableOnly(v.doc, v.tc, ns, name)
			if er != nil {
				return er
			}
			items = append(items, orderItem{n, map[string]int{"top": 1, "left": 2, "bottom": 3, "right": 4}})
		}
	}
	for _, item := range items {
		last := 0
		seen := map[string]bool{}
		for _, child := range v.doc.Elements() {
			p, ok := child.Parent()
			if !ok || p != item.n {
				continue
			}
			rank := item.rank[child.Name().Local]
			if child.Name().Space != ns || rank == 0 || rank < last || seen[child.Name().Local] {
				return fmt.Errorf("saved table child order/duplicate %s", child.Name().Local)
			}
			seen[child.Name().Local] = true
			last = rank
		}
	}
	return nil
}

// Mask only the declared owner attribute or direct child. Whole-source comparison
// retains all neighboring cells, text, metadata, nonselected attributes and bytes.
func tableMask(raw []byte, r tableRecord) ([]byte, error) {
	v, e := tableViewOf(raw, r.Format)
	if e != nil {
		return nil, e
	}
	var mask struct {
		Owner, NestedOwner   string
		Attributes, Children []string
		Multiple             []struct {
			Owner      string
			Attributes []string
		}
	}
	if e = json.Unmarshal(r.Mask, &mask); e != nil {
		return nil, e
	}
	ranges := []formattingRange{}
	addAttrs := func(n losslessxml.Element, keys []string) error {
		a, b := n.SourceRange()
		if a < 0 || b > len(raw) {
			return fmt.Errorf("mask source range")
		}
		allow := map[string]bool{}
		for _, key := range keys {
			allow[key] = true
		}
		_, list, er := formattingOpeningMasks(raw[a:b], allow)
		if er != nil {
			return er
		}
		for _, m := range list {
			ranges = append(ranges, formattingRange{start: a + m.start, end: a + m.end})
		}
		return nil
	}
	addNode := func(n losslessxml.Element) { a, b := n.SourceRange(); ranges = append(ranges, formattingRange{a, b}) }
	ns := pptxA
	if r.Format == "docx" {
		ns = retainedDocxW
	}
	owner := v.tc
	switch mask.Owner {
	case "tcPr":
		owner = v.tc
	case "tblPr":
		owner = v.tblPr
	case "trPr":
		owner = v.trPr
	case "frameOff":
		owner = v.off
	case "cellRunPr":
		owner = v.runPr
	case "":
		if len(mask.Multiple) == 0 {
			return nil, fmt.Errorf("empty edit mask")
		}
	default:
		return nil, fmt.Errorf("unknown owner %s", mask.Owner)
	}
	if mask.NestedOwner != "" {
		owner, e = tableOnly(v.doc, owner, ns, mask.NestedOwner)
		if e != nil {
			return nil, e
		}
	}
	// Nested PPTX border masks already name the line's w attribute and fill
	// children; the generic path below preserves its unpatched dash bytes.
	if r.Format == "docx" && r.Kind == "cell-margins" {
		// The sealed margin mask names side children; only w:w may change.
		// The unpatched w:type must retain its quote style and literal bytes.
		for _, side := range []string{"top", "left", "bottom", "right"} {
			n, er := tableOnly(v.doc, owner, ns, side)
			if er != nil {
				return nil, er
			}
			if er = addAttrs(n, []string{"w:w"}); er != nil {
				return nil, er
			}
		}
	} else if r.Format == "docx" && (strings.HasPrefix(r.Kind, "border-") && r.Kind != "border-remove" || r.Kind == "cell-anchor" || r.Kind == "cell-direction" || r.Kind == "cell-no-wrap" || r.Kind == "cell-fit-text" || r.Kind == "cell-hide-mark" || r.Kind == "row-header" || r.Kind == "row-cant-split" || r.Kind == "row-height" || r.Kind == "table-alignment") {
		keys := []string{"w:val"}
		if strings.HasPrefix(r.Kind, "border-") {
			keys = []string{"w:val", "w:sz", "w:color"}
		}
		if r.Kind == "row-height" {
			keys = []string{"w:val", "w:hRule"}
		}
		if len(mask.Children) != 1 {
			return nil, fmt.Errorf("scalar child mask")
		}
		n, er := tableOnly(v.doc, owner, ns, mask.Children[0])
		if er != nil {
			return nil, er
		}
		if er = addAttrs(n, keys); er != nil {
			return nil, er
		}
	} else {
		for _, name := range mask.Children {
			nodes := retainedChild(v.doc, owner, ns, name)
			if len(nodes) > 1 {
				return nil, fmt.Errorf("ambiguous masked child")
			}
			if len(nodes) == 1 {
				addNode(nodes[0])
			}
		}
	}
	if r.Format == "docx" {
		for i, k := range mask.Attributes {
			mask.Attributes[i] = "w:" + k
		}
	}
	if e = addAttrs(owner, mask.Attributes); e != nil {
		return nil, e
	}
	for _, m := range mask.Multiple {
		var n losslessxml.Element
		switch m.Owner {
		case "gridCol0":
			cols := retainedChild(v.doc, v.grid, ns, "gridCol")
			if len(cols) != 3 {
				return nil, fmt.Errorf("grid columns")
			}
			n = cols[0]
		case "row0":
			n = v.row
		case "frameExt":
			n = v.ext
		default:
			return nil, fmt.Errorf("unknown span %s", m.Owner)
		}
		if e = addAttrs(n, m.Attributes); e != nil {
			return nil, e
		}
	}
	return formattingStrip(raw, ranges)
}
func tableMaskedSame(before, after []byte, r tableRecord) error {
	left, e := tableMask(before, r)
	if e != nil {
		return e
	}
	right, e := tableMask(after, r)
	if e != nil {
		return e
	}
	if !bytes.Equal(left, right) {
		return fmt.Errorf("unpatched table bytes drift %s", r.ID)
	}
	return nil
}
