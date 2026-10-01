package spreadsheet

import (
	"encoding/xml"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// SetInlineStringWrap edits one existing inline-string cell and installs a
// separate direct wrap-text XF copied from the cell's current simple XF.
// Both XML parts are preflighted and committed as one graph plan; the legacy
// workbook writer and other cell styles are untouched.
func (s *EditSession) SetInlineStringWrap(sheet, address, value string, wrap bool) error {
	if s == nil || s.pkg == nil || !utf8.ValidString(value) {
		return editRefusal("unsupported_structure", "invalid inline-string edit")
	}
	ref, err := utils.ParseCellRef(address)
	if err != nil || !strictCell.MatchString(address) || ref.Row > 1048576 || ref.Col > 16384 {
		return editRefusal("unsupported_structure", "expected bounded A1 address")
	}
	part := s.sheets[sheet]
	if part == "" {
		return editRefusal("missing_target", "sheet absent")
	}
	if err := s.guard(); err != nil {
		return err
	}
	data, hash, err := s.pkg.Part(part)
	if err != nil {
		return err
	}
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return editRefusal("unsupported_structure", "invalid worksheet XML")
	}
	var target, text losslessxml.Element
	matches := 0
	for _, node := range doc.Elements() {
		if node.Name() == expanded("c") && attr(node, "r") == address {
			matches++
			target = node
		}
	}
	if matches != 1 || attr(target, "t") != "inlineStr" {
		return editRefusal("unsupported_structure", "exact existing inline-string cell required")
	}
	styleIndex := 0
	if raw := attr(target, "s"); raw != "" {
		styleIndex, err = strconv.Atoi(raw)
		if err != nil || styleIndex < 0 {
			return editRefusal("unsupported_structure", "invalid cell style index")
		}
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	stylesPart := ""
	for _, edge := range graph.Edges {
		if edge.Source == s.main && edge.Type == packaging.RelTypeStyles {
			if stylesPart != "" || edge.External {
				return editRefusal("unsupported_structure", "ambiguous styles relationship")
			}
			stylesPart = edge.ResolvedPart
		}
	}
	if stylesPart == "" {
		return editRefusal("unsupported_structure", "styles part required for direct wrap XF")
	}
	stylesBytes, stylesHash, err := s.pkg.Part(stylesPart)
	if err != nil {
		return err
	}
	stylesDoc, err := losslessxml.Parse(stylesBytes)
	if err != nil {
		return editRefusal("unsupported_structure", "invalid styles XML")
	}
	var table, base losslessxml.Element
	for _, node := range stylesDoc.Elements() {
		if node.Name() != expanded("cellXfs") {
			continue
		}
		parent, ok := node.Parent()
		if !ok || parent != stylesDoc.Elements()[0] || table.Ordinal() >= 0 {
			return editRefusal("unsupported_structure", "ambiguous cellXfs")
		}
		table = node
	}
	if table.Ordinal() < 0 {
		return editRefusal("unsupported_structure", "missing cellXfs")
	}
	count := 0
	for _, node := range stylesDoc.Elements() {
		parent, ok := node.Parent()
		if !ok || parent != table {
			continue
		}
		if node.Name() != expanded("xf") {
			return editRefusal("unsupported_structure", "unknown cellXfs child")
		}
		if count == styleIndex {
			base = node
		}
		count++
	}
	if count < 1 || count >= 65535 || styleIndex >= count {
		return editRefusal("unsupported_structure", "invalid cellXfs/style index")
	}
	baseAttrs := base.Attributes()
	for i := range baseAttrs {
		if baseAttrs[i].Name.Space != "" {
			return editRefusal("unsupported_structure", "unsupported XF attribute namespace")
		}
		if baseAttrs[i].Name.Local == "applyAlignment" {
			baseAttrs[i].Value = "1"
		}
	}
	foundApply := false
	for _, a := range baseAttrs {
		if a.Name.Local == "applyAlignment" {
			foundApply = true
		}
	}
	if !foundApply {
		baseAttrs = append(baseAttrs, xml.Attr{Name: xml.Name{Local: "applyAlignment"}, Value: "1"})
	}
	alignment := losslessxml.NewElement{Name: expanded("alignment"), Attributes: []xml.Attr{{Name: xml.Name{Local: "wrapText"}, Value: strconv.FormatBool(wrap)}}}
	seenAlignment := false
	for _, node := range stylesDoc.Elements() {
		parent, ok := node.Parent()
		if !ok || parent != base {
			continue
		}
		if node.Name() != expanded("alignment") || seenAlignment {
			return editRefusal("unsupported_structure", "XF contains unsupported nested properties")
		}
		seenAlignment = true
		attrs := node.Attributes()
		for i := range attrs {
			if attrs[i].Name.Space != "" {
				return editRefusal("unsupported_structure", "unsupported alignment attribute namespace")
			}
			if attrs[i].Name.Local == "wrapText" {
				attrs[i].Value = strconv.FormatBool(wrap)
			}
		}
		wrapPresent := false
		for _, a := range attrs {
			if a.Name.Local == "wrapText" {
				wrapPresent = true
			}
		}
		if !wrapPresent {
			attrs = append(attrs, xml.Attr{Name: xml.Name{Local: "wrapText"}, Value: strconv.FormatBool(wrap)})
		}
		alignment.Attributes = attrs
	}
	child := losslessxml.NewElement{Name: expanded("xf"), Attributes: baseAttrs, Children: []losslessxml.NewElement{alignment}}
	editedStyles, err := stylesDoc.InsertChildren([]losslessxml.ChildInsertion{{Parent: table, Children: []losslessxml.NewElement{child}}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	stylesDoc, err = losslessxml.Parse(editedStyles)
	if err != nil {
		return err
	}
	for _, node := range stylesDoc.Elements() {
		if node.Name() != expanded("cellXfs") {
			continue
		}
		parent, ok := node.Parent()
		if ok && parent == stylesDoc.Elements()[0] {
			editedStyles, err = stylesDoc.Edit(nil, []losslessxml.AttributeEdit{{Target: node, Name: xml.Name{Local: "count"}, Value: strconv.Itoa(count + 1)}})
			break
		}
	}
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	row, ok := target.Parent()
	if !ok || row.Name() != expanded("row") {
		return editRefusal("unsupported_structure", "invalid row owner")
	}
	sheetData, ok := row.Parent()
	if !ok || sheetData.Name() != expanded("sheetData") {
		return editRefusal("unsupported_structure", "invalid sheetData owner")
	}
	root, ok := sheetData.Parent()
	if !ok || root != doc.Elements()[0] {
		return editRefusal("unsupported_structure", "invalid worksheet owner")
	}
	inline := 0
	for _, node := range doc.Elements() {
		parent, ok := node.Parent()
		if ok && parent == target {
			if node.Name() != expanded("is") {
				return editRefusal("unsupported_structure", "mixed cell content")
			}
			inline++
			for _, child := range doc.Elements() {
				owner, ok := child.Parent()
				if ok && owner == node {
					if child.Name() != expanded("t") || text.Ordinal() >= 0 {
						return editRefusal("unsupported_structure", "mixed inline string")
					}
					text = child
				}
			}
		}
	}
	if inline != 1 || text.Ordinal() < 0 {
		return editRefusal("unsupported_structure", "one inline text leaf required")
	}
	if _, leaf := text.Text(); !leaf || strings.ContainsRune(value, '\x00') {
		return editRefusal("unsupported_structure", "invalid inline text")
	}
	editedSheet, err := doc.Edit([]losslessxml.TextEdit{{Target: text, Text: value}}, []losslessxml.AttributeEdit{{Target: target, Name: xml.Name{Local: "s"}, Value: strconv.Itoa(count)}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	plan, err := s.pkg.PlanGraphMutation(packaging.GraphMutation{Replacements: []packaging.Replacement{{Part: stylesPart, ExpectedSHA256: stylesHash, Data: editedStyles}, {Part: part, ExpectedSHA256: hash, Data: editedSheet}}})
	if err != nil {
		return err
	}
	if err := s.pkg.ApplyGraphPlan(plan); err != nil {
		return err
	}
	s.generation++
	return nil
}
