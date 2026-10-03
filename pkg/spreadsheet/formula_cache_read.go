package spreadsheet

import (
	"math"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// ReadFormulaCache returns the stored numeric result of one ordinary formula,
// or nil when no cache exists. It does not evaluate the formula or edit a part.
// Other result types and attributed/shared/array topologies are not inferred.
func (s *EditSession) ReadFormulaCache(sheet, address string) (*float64, error) {
	ref, err := utils.ParseCellRef(address)
	if err != nil || !strictCell.MatchString(address) || ref.Row > 1048576 || ref.Col > 16384 {
		return nil, editRefusal("unsupported_structure", "exact bounded formula cell address required")
	}
	part := s.sheets[sheet]
	if part == "" {
		return nil, editRefusal("missing_target", "sheet absent")
	}
	body, _, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	doc, err := losslessxml.Parse(body)
	if err != nil {
		return nil, err
	}
	if doc.Elements()[0].Name() != expanded("worksheet") {
		return nil, editRefusal("unsupported_structure", "worksheet root required")
	}
	var cell losslessxml.Element
	count := 0
	for _, node := range doc.Elements() {
		if node.Name() == expanded("c") && attr(node, "r") == address {
			cell = node
			count++
		}
	}
	if count != 1 {
		if count == 0 {
			return nil, editRefusal("missing_target", "formula cell absent")
		}
		return nil, editRefusal("ambiguous_target", "formula cell duplicated")
	}
	row, ok := cell.Parent()
	if !ok || row.Name() != expanded("row") || attr(row, "r") != strconv.Itoa(ref.Row) {
		return nil, editRefusal("unsupported_structure", "formula row differs")
	}
	container, ok := row.Parent()
	if !ok || container.Name() != expanded("sheetData") {
		return nil, editRefusal("unsupported_structure", "formula container differs")
	}
	owner, ok := container.Parent()
	if !ok || owner != doc.Elements()[0] {
		return nil, editRefusal("unsupported_structure", "formula sheet differs")
	}
	for _, a := range cell.Attributes() {
		if a.Name.Space != "" || a.Name.Local != "r" && a.Name.Local != "s" {
			return nil, editRefusal("unsupported_structure", "unclassified formula cell attribute")
		}
	}
	var formula, cache losslessxml.Element
	for _, node := range doc.Elements() {
		parent, ok := node.Parent()
		if !ok || parent != cell {
			continue
		}
		switch node.Name() {
		case expanded("f"):
			if formula.Ordinal() >= 0 {
				return nil, editRefusal("unsupported_structure", "duplicate formula")
			}
			formula = node
		case expanded("v"):
			if cache.Ordinal() >= 0 {
				return nil, editRefusal("unsupported_structure", "duplicate formula cache")
			}
			cache = node
		default:
			return nil, editRefusal("unsupported_structure", "unclassified formula child")
		}
	}
	if formula.Ordinal() < 0 {
		return nil, editRefusal("unsupported_structure", "selected cell has no formula")
	}
	if len(formula.Attributes()) != 0 {
		return nil, editRefusal("xlsx-cache-topology-unsupported", "attributed formula cache read")
	}
	if _, leaf := formula.Text(); !leaf {
		return nil, editRefusal("unsupported_structure", "nonleaf formula")
	}
	if cache.Ordinal() < 0 {
		return nil, nil
	}
	if len(cache.Attributes()) != 0 {
		return nil, editRefusal("unsupported_structure", "attributed formula value")
	}
	text, leaf := cache.Text()
	value, err := strconv.ParseFloat(text, 64)
	if !leaf || err != nil || math.IsNaN(value) || math.IsInf(value, 0) {
		return nil, editRefusal("unsupported_structure", "non-finite or unsupported formula cache")
	}
	return &value, nil
}
