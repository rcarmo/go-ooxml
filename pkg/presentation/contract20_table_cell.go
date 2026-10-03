package presentation

import (
	"bytes"
	"encoding/xml"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ContractTableCellTarget is an issued, part-scoped cell selector. Ordinary
// table validity is checked at edit preflight, never inferred from ZIP intake.
type ContractTableCellTarget struct{ selected *RetainedTableTarget }

func (s *EditSession) SetContractTableCellText(target *ContractTableCellTarget, text string) error {
	if target == nil || target.selected == nil {
		return tableRefuse("PPTX_STALE_TABLE_HANDLE", "missing table cell target")
	}
	if _, issued := s.contractTableCells[target]; !issued {
		return tableRefuse("PPTX_STALE_TABLE_HANDLE", "table cell handle was not issued by this session")
	}
	t := target.selected
	// Contract-owned handles classify actual session/generation/part-fingerprint
	// currency directly. The legacy RetainedTableTarget.tableCurrent contract
	// remains unchanged and continues returning stale_target for legacy edits.
	if t.session != s || t.generation != s.generation {
		return tableRefuse("PPTX_STALE_TABLE_HANDLE", "foreign or stale table cell target")
	}
	_, currentHash, err := s.pkg.Part(t.part)
	if err != nil {
		return err
	}
	if currentHash != t.hash {
		return tableRefuse("PPTX_STALE_TABLE_HANDLE", "table part changed")
	}
	if _, err := losslessxml.Parse([]byte("<t>" + xmlEscapeContractText(text) + "</t>")); err != nil {
		return tableRefuse("PPTX_ARGUMENT_INVALID", "invalid XML cell text")
	}
	rows := tableChildren(t.doc, t.table, packaging.NSDrawingML, "tr")
	grid, err := tableOne(t.doc, t.table, packaging.NSDrawingML, "tblGrid")
	if err != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "table grid absent")
	}
	cols := tableChildren(t.doc, grid, packaging.NSDrawingML, "gridCol")
	if len(rows) == 0 || len(cols) == 0 {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "empty table")
	}
	merged := false
	for _, row := range rows {
		occupancy := 0
		for _, cell := range tableChildren(t.doc, row, packaging.NSDrawingML, "tc") {
			width := 1
			for _, a := range cell.Attributes() {
				if a.Name.Space != "" {
					return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "foreign cell attribute")
				}
				switch a.Name.Local {
				case "gridSpan", "rowSpan":
					n, e := strictPositiveContractSpan(a.Value)
					if e != nil {
						return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid table span")
					}
					if n > 1 {
						merged = true
					}
					if a.Name.Local == "gridSpan" {
						width = n
					}
				case "hMerge", "vMerge":
					if a.Value != "0" && a.Value != "1" && a.Value != "true" && a.Value != "false" {
						return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid merge flag")
					}
					if a.Value == "1" || a.Value == "true" {
						merged = true
					}
				default:
					return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "unknown cell attribute")
				}
			}
			occupancy += width
		}
		if occupancy != len(cols) {
			return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "logical row occupancy mismatch")
		}
	}
	if merged {
		return tableRefuse("PPTX_TABLE_MERGE_UNSUPPORTED", "merged table text edit")
	}
	// This text-only profile does not require legacy table-property locks or
	// tableStyleId. It does require direct table ownership and rectangular,
	// bounded geometry; merged/invalid topology was classified above.
	if err := contractTableTextAdmit(t.doc, t.frame, t.table, rows, cols); err != nil {
		return err
	}
	body, err := tableOne(t.doc, t.cell, packaging.NSDrawingML, "txBody")
	if err != nil {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "cell body absent")
	}
	paras := tableChildren(t.doc, body, packaging.NSDrawingML, "p")
	if len(paras) != 1 {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "single paragraph required")
	}
	// Only the exact empty direct paragraph forms have no text leaf. Author a
	// single ordinary run there; do not infer an editable target from any other
	// paragraph structure or reconstruct a caller-owned slide.
	plain := paras[0].Raw()
	if bytes.Equal(plain, []byte("<a:p></a:p>")) || bytes.Equal(plain, []byte("<a:p/>")) {
		if text == "" {
			return nil
		}
		leaf := losslessxml.NewElement{Name: name(packaging.NSDrawingML, "t"), Text: text}
		if strings.TrimSpace(text) != text {
			leaf.Attributes = []xml.Attr{{Name: xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}, Value: "preserve"}}
		}
		data, e := t.doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: paras[0], Children: []losslessxml.NewElement{{Name: name(packaging.NSDrawingML, "r"), Children: []losslessxml.NewElement{leaf}}}}})
		if e != nil {
			return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", e.Error())
		}
		if e = s.pkg.Replace([]packaging.Replacement{{Part: t.part, ExpectedSHA256: t.hash, Data: data}}); e != nil {
			return e
		}
		s.generation++
		return nil
	}
	var leaf losslessxml.Element
	count := 0
	for _, node := range t.doc.Elements() {
		if node.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: "t"}) {
			continue
		}
		a, b := node.SourceRange()
		start, end := t.cell.SourceRange()
		if a >= start && b <= end {
			leaf = node
			count++
		}
	}
	if count != 1 {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "single visible run required")
	}
	run, ok := leaf.Parent()
	if !ok || run.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: "r"}) {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "direct run required")
	}
	p, ok := run.Parent()
	if !ok || p != paras[0] {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "owned paragraph required")
	}
	raw := leaf.Raw()
	if len(raw) < 4 || strings.HasSuffix(string(raw), "/>") {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", "text leaf required")
	}
	data, err := t.doc.ReplaceText([]losslessxml.TextEdit{{Target: leaf, Text: text}})
	if err != nil {
		return tableRefuse("PPTX_UNSUPPORTED_TEXT_TOPOLOGY", err.Error())
	}
	if err := s.pkg.Replace([]packaging.Replacement{{Part: t.part, ExpectedSHA256: t.hash, Data: data}}); err != nil {
		return err
	}
	s.generation++
	return nil
}

func contractTableTextAdmit(d *losslessxml.Document, frame, tbl losslessxml.Element, rows, cols []losslessxml.Element) error {
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	if tableDirectProfile(d, frame, []xml.Name{p("nvGraphicFramePr"), p("xfrm"), a("graphic")}, xml.Name{}) != nil ||
		tableDirectProfile(d, tbl, []xml.Name{a("tblPr"), a("tblGrid")}, a("tr")) != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "unknown frame or table owner")
	}
	nv, err := tableOne(d, frame, packaging.NSPresentationML, "nvGraphicFramePr")
	if err != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "missing frame identity")
	}
	if tableDirectProfile(d, nv, []xml.Name{p("cNvPr"), p("cNvGraphicFramePr"), p("nvPr")}, xml.Name{}) != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid frame identity")
	}
	graphic, err := tableOne(d, frame, packaging.NSDrawingML, "graphic")
	if err != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "missing graphic")
	}
	data, err := tableOne(d, graphic, packaging.NSDrawingML, "graphicData")
	if err != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "missing graphic data")
	}
	if uri, ok := tableAttr(data, "uri"); !ok || uri != "http://schemas.openxmlformats.org/drawingml/2006/table" {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "foreign graphic data")
	}
	if tableDirectProfile(d, data, []xml.Name{a("tbl")}, xml.Name{}) != nil {
		return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "foreign table payload")
	}
	for _, col := range cols {
		if _, err := tableNumber(col, "w", 1, 100000000); err != nil {
			return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid grid width")
		}
	}
	for _, row := range rows {
		if _, err := tableNumber(row, "h", 1, 100000000); err != nil {
			return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid row height")
		}
		if len(tableChildren(d, row, packaging.NSDrawingML, "tc")) != len(cols) {
			return tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "nonrectangular ordinary table")
		}
	}
	return nil
}

// Isolate the XML character preflight from the source-text replacement. This
// helper escapes only the five XML metacharacters; Parse validates codepoints.
func xmlEscapeContractText(value string) string {
	r := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", "\"", "&quot;", "'", "&apos;")
	return r.Replace(value)
}
func strictPositiveContractSpan(text string) (int, error) {
	if text == "" || len(text) > 8 {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid span")
	}
	n := 0
	for _, r := range text {
		if r < '0' || r > '9' {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid span")
		}
		n = n*10 + int(r-'0')
	}
	if n < 1 {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid span")
	}
	return n, nil
}
