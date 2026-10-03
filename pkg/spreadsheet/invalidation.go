package spreadsheet

import (
	"encoding/xml"
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/formula"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// CalculationEffect reports an applied numeric edit and its cache effects.
// It is not proof of filesystem delivery or numeric calculation.
type CalculationEffect struct {
	State                   string   `json:"state"`
	Invalidated             []string `json:"invalidated"`
	ValueChanged            bool     `json:"value_changed"`
	RemovedCalculationChain string   `json:"removed_calculation_chain,omitempty"`
}
type dependencyRange struct {
	part                       string
	loRow, hiRow, loCol, hiCol int
}
type formulaCell struct {
	part, sheet, address string
	row, col             int
	cache                losslessxml.Element
	hasCache             bool
	deps                 []dependencyRange
}
type formulaSheet struct {
	doc   *losslessxml.Document
	hash  string
	cells map[string]losslessxml.Element
}

// SetNumberWithInvalidation is opt-in: the entire workbook must fit the static
// reference subset. Affected caches clear in the same multi-part transaction as
// the input and recalculation metadata. No formulas are evaluated.
func (s *EditSession) SetNumberWithInvalidation(target *NumberTarget, value float64) (CalculationEffect, error) {
	return s.setNumberWithInvalidation(target, value, false, false)
}

// SetNumberWithAllCachesInvalidated is an opt-in retained-edit profile: after
// a changed numeric input, every ordinary formula cache is invalidated, even
// when the formula is independent of that input. It never evaluates formulas.
// The existing dependency-scoped method retains its original contract.
func (s *EditSession) SetNumberWithAllCachesInvalidated(target *NumberTarget, value float64) (CalculationEffect, error) {
	return s.setNumberWithInvalidation(target, value, true, false)
}

// SetNumberWithSealedOpaqueCaches retains only the exact, inert Contract20
// root cache edges. It is not permission for chart/external dependency edits.
func (s *EditSession) SetNumberWithSealedOpaqueCaches(target *NumberTarget, value float64) (CalculationEffect, error) {
	if err := s.sealedOpaqueCachePreflight(); err != nil {
		return CalculationEffect{State: "unchanged", Invalidated: []string{}}, err
	}
	return s.setNumberWithInvalidation(target, value, true, true)
}

func (s *EditSession) setNumberWithInvalidation(target *NumberTarget, value float64, allCaches, sealedOpaque bool) (CalculationEffect, error) {
	result := CalculationEffect{State: "unchanged", Invalidated: []string{}}
	if target == nil || target.session != s || target.generation != s.generation || target.consumed {
		return result, editRefusal("stale_target", "stale, foreign or consumed numeric target")
	}
	if math.IsNaN(value) || math.IsInf(value, 0) {
		return result, editRefusal("unsupported_structure", "finite value required")
	}
	_, current, err := s.pkg.Part(target.part)
	if err != nil {
		return result, err
	}
	if current != target.hash {
		return result, editRefusal("stale_target", "numeric source changed")
	}
	// A no-op on this held numeric target cannot change any formula cache.
	// This is deliberately before attributed-formula topology preflight; a
	// changed input still traverses that preflight and refuses atomically.
	if allCaches {
		old, parseErr := strconv.ParseFloat(target.text, 64)
		if parseErr != nil {
			return result, editRefusal("stale_target", "invalid held numeric value")
		}
		if old == value {
			return result, nil
		}
	}
	if err = s.guardModeWithOpaque(true, sealedOpaque); err != nil {
		return result, err
	}
	mainData, mainHash, err := s.pkg.Part(s.main)
	if err != nil {
		return result, err
	}
	main, err := losslessxml.Parse(mainData)
	if err != nil {
		return result, err
	}
	var calc losslessxml.Element
	calcCount := 0
	for _, e := range main.Elements() {
		if e.Name() != expanded("calcPr") {
			continue
		}
		parent, ok := e.Parent()
		if !ok || parent != main.Elements()[0] {
			return result, editRefusal("unsupported_structure", "unexpected calculation metadata owner")
		}
		calc = e
		calcCount++
		for _, a := range e.Attributes() {
			if a.Name.Space != "" {
				return result, editRefusal("unsupported_structure", "unknown calc metadata namespace")
			}
			switch a.Name.Local {
			case "calcMode":
				if a.Value != "auto" {
					return result, editRefusal("unsupported_structure", "non-auto calculation mode")
				}
			case "iterate":
				if a.Value != "0" && a.Value != "false" {
					return result, editRefusal("unsupported_structure", "iterative calculation not modelled")
				}
			case "fullPrecision":
				if a.Value != "1" && a.Value != "true" {
					return result, editRefusal("unsupported_structure", "precision-as-displayed dependencies not modelled")
				}
			case "calcId", "fullCalcOnLoad", "forceFullCalc", "calcOnSave", "concurrentCalc", "concurrentManualCount":
			default:
				return result, editRefusal("unsupported_structure", "unknown calculation metadata")
			}
		}
	}
	if calcCount > 1 {
		return result, editRefusal("unsupported_structure", "duplicate calcPr")
	}
	sheetNames := make([]string, 0, len(s.sheets))
	for name := range s.sheets {
		sheetNames = append(sheetNames, name)
	}
	sort.Strings(sheetNames)
	lookup := map[string]string{}
	for _, name := range sheetNames {
		key := strings.ToLower(name)
		if _, ok := lookup[key]; ok {
			return result, editRefusal("ambiguous_target", "case-colliding sheet names")
		}
		lookup[key] = s.sheets[name]
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return result, err
	}
	enrolled := map[string]bool{}
	for _, p := range s.sheets {
		enrolled[p] = true
	}
	for _, p := range graph.Parts {
		if p.ContentType == packaging.ContentTypeWorksheet && !enrolled[p.Name] {
			return result, editRefusal("unsupported_structure", "unlisted worksheet could contain dependencies")
		}
	}
	sheets := map[string]formulaSheet{}
	formulas := []formulaCell{}
	for _, sheetName := range sheetNames {
		part := s.sheets[sheetName]
		data, hash, err := s.pkg.Part(part)
		if err != nil {
			return result, err
		}
		doc, err := losslessxml.Parse(data)
		if err != nil {
			return result, err
		}
		es := doc.Elements()
		if es[0].Name() != expanded("worksheet") {
			return result, editRefusal("unsupported_structure", "worksheet root required")
		}
		sheetDataCount := 0
		rowIDs := map[string]bool{}
		for _, e := range es {
			if e.Name() == expanded("sheetData") {
				p, ok := e.Parent()
				if !ok || p != es[0] {
					return result, editRefusal("unsupported_structure", "nested sheetData")
				}
				sheetDataCount++
			}
			if e.Name() == expanded("row") {
				id := attr(e, "r")
				if rowIDs[id] {
					return result, editRefusal("ambiguous_target", "duplicate row identity")
				}
				rowIDs[id] = true
			}
		}
		if sheetDataCount != 1 {
			return result, editRefusal("unsupported_structure", "exactly one sheetData container required")
		}
		cells := map[string]losslessxml.Element{}
		children := map[losslessxml.Element][]losslessxml.Element{}
		for _, e := range es {
			if p, ok := e.Parent(); ok {
				children[p] = append(children[p], e)
			}
		}
		for _, e := range es {
			if e.Name() != expanded("c") {
				continue
			}
			address := attr(e, "r")
			coordinate, err := formula.ParseCell(address)
			if err != nil || !strictCell.MatchString(address) {
				return result, editRefusal("unsupported_structure", "noncanonical cell address")
			}
			if _, ok := cells[address]; ok {
				return result, editRefusal("ambiguous_target", "duplicate cell address")
			}
			cells[address] = e
			row, ok := e.Parent()
			if !ok || row.Name() != expanded("row") {
				return result, editRefusal("unsupported_structure", "cell without row")
			}
			rowNumber, err := unsignedIndex(attr(row, "r"))
			if err != nil || rowNumber != uint64(coordinate.Row) {
				return result, editRefusal("unsupported_structure", "row/cell address mismatch")
			}
			sd, ok := row.Parent()
			if !ok || sd.Name() != expanded("sheetData") {
				return result, editRefusal("unsupported_structure", "row without sheetData")
			}
			root, ok := sd.Parent()
			if !ok || root != es[0] {
				return result, editRefusal("unsupported_structure", "unexpected worksheet owner")
			}
			var f, v losslessxml.Element
			fc, vc := 0, 0
			for _, child := range children[e] {
				switch child.Name() {
				case expanded("f"):
					f = child
					fc++
				case expanded("v"):
					v = child
					vc++
				}
			}
			for _, a := range e.Attributes() {
				if a.Name.Space != "" || a.Name.Local != "r" && a.Name.Local != "s" && a.Name.Local != "t" {
					return result, editRefusal("unsupported_structure", "unclassified cell metadata")
				}
			}
			if vc > 1 {
				return result, editRefusal("unsupported_structure", "duplicate cell value")
			}
			if vc == 1 {
				_, leaf := v.Text()
				if !leaf || len(v.Attributes()) != 0 {
					return result, editRefusal("unsupported_structure", "invalid value leaf")
				}
			}
			if fc == 0 {
				continue
			}
			if fc != 1 || vc > 1 {
				return result, editRefusal("unsupported_structure", "duplicate formula/cache")
			}
			if len(f.Attributes()) != 0 {
				if allCaches && (attr(f, "t") == "shared" || attr(f, "t") == "array" || attr(f, "t") == "dataTable") {
					return result, editRefusal("xlsx-cache-topology-unsupported", "attributed shared/array/dataTable result range")
				}
				return result, editRefusal("unsupported_structure", "shared/array/formula metadata not supported")
			}
			for _, child := range children[e] {
				if child != f && child != v {
					return result, editRefusal("unsupported_structure", "unknown formula-cell structure")
				}
			}
			for _, a := range e.Attributes() {
				if a.Name.Space != "" || a.Name.Local != "r" && a.Name.Local != "s" && a.Name.Local != "t" {
					return result, editRefusal("unsupported_structure", "unknown formula-cell metadata")
				}
			}
			kind := attr(e, "t")
			if kind != "" && kind != "n" && kind != "str" && kind != "b" && kind != "e" {
				return result, editRefusal("unsupported_structure", "unsupported formula result type")
			}
			expression, leaf := f.Text()
			if !leaf {
				return result, editRefusal("unsupported_structure", "mixed formula text")
			}
			references, err := formula.Analyze(expression)
			if err != nil {
				return result, editRefusal("unsupported_structure", err.Error())
			}
			node := formulaCell{part: part, sheet: sheetName, address: address, row: coordinate.Row, col: coordinate.Column, cache: v, hasCache: vc == 1}
			for _, ref := range references {
				dependencyPart := part
				if ref.Sheet != "" {
					var ok bool
					dependencyPart, ok = lookup[strings.ToLower(ref.Sheet)]
					if !ok {
						return result, editRefusal("missing_target", "formula references absent worksheet")
					}
				}
				node.deps = append(node.deps, dependencyRange{dependencyPart, min(ref.First.Row, ref.Last.Row), max(ref.First.Row, ref.Last.Row), min(ref.First.Column, ref.Last.Column), max(ref.First.Column, ref.Last.Column)})
			}
			formulas = append(formulas, node)
		}
		// Formula elements outside cells are not silently omitted from the graph.
		for _, e := range es {
			if e.Name() == expanded("f") {
				parent, ok := e.Parent()
				if !ok || parent.Name() != expanded("c") {
					return result, editRefusal("unsupported_structure", "formula outside cell")
				}
			}
		}
		sheets[part] = formulaSheet{doc, hash, cells}
	}
	old, _ := strconv.ParseFloat(target.text, 64)
	if old == value {
		return result, nil
	}
	owner, ok := target.element.Parent()
	if !ok {
		return result, editRefusal("stale_target", "numeric owner absent")
	}
	address := attr(owner, "r")
	coord, err := formula.ParseCell(address)
	if err != nil {
		return result, err
	}
	type changedCell struct {
		part     string
		row, col int
	}
	dirty := []changedCell{{target.part, coord.Row, coord.Column}}
	affected := map[int]bool{}
	if allCaches {
		for i := range formulas {
			affected[i] = true
		}
	}
	// Finite fixed-point propagation supports cycles without looping forever.
	for cursor := 0; cursor < len(dirty); cursor++ {
		c := dirty[cursor]
		for i, f := range formulas {
			if affected[i] {
				continue
			}
			for _, dep := range f.deps {
				if c.part == dep.part && c.row >= dep.loRow && c.row <= dep.hiRow && c.col >= dep.loCol && c.col <= dep.hiCol {
					affected[i] = true
					dirty = append(dirty, changedCell{f.part, f.row, f.col})
					break
				}
			}
		}
	}
	edits := map[string][]losslessxml.TextEdit{}
	cacheRemovals := map[string][]losslessxml.Element{}
	inputSheet := sheets[target.part]
	inputCell, ok := inputSheet.cells[address]
	if !ok {
		return result, editRefusal("stale_target", "input cell absent")
	}
	var inputValue losslessxml.Element
	for _, e := range inputSheet.doc.Elements() {
		if p, ok := e.Parent(); ok && p == inputCell && e.Name() == expanded("v") {
			inputValue = e
		}
	}
	if inputValue.Name() != expanded("v") {
		return result, editRefusal("stale_target", "input value absent")
	}
	edits[target.part] = append(edits[target.part], losslessxml.TextEdit{Target: inputValue, Text: strconv.FormatFloat(value, 'g', -1, 64)})
	for i, f := range formulas {
		if !affected[i] {
			continue
		}
		result.Invalidated = append(result.Invalidated, f.sheet+"!"+f.address)
		if f.hasCache {
			cache, leaf := f.cache.Text()
			if !leaf || len(f.cache.Attributes()) != 0 {
				return CalculationEffect{}, editRefusal("unsupported_structure", "unsupported cached-value structure")
			}
			if allCaches {
				cacheRemovals[f.part] = append(cacheRemovals[f.part], f.cache)
			} else if cache != "" {
				edits[f.part] = append(edits[f.part], losslessxml.TextEdit{Target: f.cache, Text: ""})
			}
		}
	}
	sort.Strings(result.Invalidated)
	replacements := []packaging.Replacement{}
	parts := []string{}
	partSet := map[string]bool{}
	for part := range edits {
		partSet[part] = true
	}
	for part := range cacheRemovals {
		partSet[part] = true
	}
	for part := range partSet {
		parts = append(parts, part)
	}
	sort.Strings(parts)
	for _, part := range parts {
		sheet := sheets[part]
		var data []byte
		if allCaches && len(cacheRemovals[part]) != 0 {
			data, err = sheet.doc.RemoveElements(cacheRemovals[part])
			if err == nil && part == target.part {
				// Re-select the input leaf from the private post-removal snapshot.
				// No package mutation occurs until every sheet and graph plan validates.
				var current *losslessxml.Document
				current, err = losslessxml.Parse(data)
				if err == nil {
					var valueLeaf losslessxml.Element
					count := 0
					for _, e := range current.Elements() {
						if e.Name() != expanded("v") {
							continue
						}
						cell, ok := e.Parent()
						if ok && cell.Name() == expanded("c") && attr(cell, "r") == address {
							valueLeaf = e
							count++
						}
					}
					if count != 1 {
						return CalculationEffect{}, editRefusal("stale_target", "input leaf missing after cache removal")
					}
					data, err = current.Edit([]losslessxml.TextEdit{{Target: valueLeaf, Text: strconv.FormatFloat(value, 'g', -1, 64)}}, nil)
				}
			}
		} else {
			data, err = sheet.doc.Edit(edits[part], nil)
		}
		if err != nil {
			return CalculationEffect{}, editRefusal("unsupported_structure", err.Error())
		}
		replacements = append(replacements, packaging.Replacement{Part: part, ExpectedSHA256: sheet.hash, Data: data})
	}
	if len(affected) > 0 {
		flags := []xml.Attr{{Name: xml.Name{Local: "calcMode"}, Value: "auto"}, {Name: xml.Name{Local: "fullCalcOnLoad"}, Value: "1"}, {Name: xml.Name{Local: "forceFullCalc"}, Value: "1"}}
		var data []byte
		if calcCount == 1 {
			attributes := []losslessxml.AttributeEdit{}
			for _, a := range flags {
				attributes = append(attributes, losslessxml.AttributeEdit{Target: calc, Name: a.Name, Value: a.Value})
			}
			data, err = main.Edit(nil, attributes)
		} else {
			data, err = main.InsertChildren([]losslessxml.ChildInsertion{{Parent: main.Elements()[0], Children: []losslessxml.NewElement{{Name: expanded("calcPr"), Attributes: flags}}}})
		}
		if err != nil {
			return CalculationEffect{}, editRefusal("unsupported_structure", err.Error())
		}
		replacements = append(replacements, packaging.Replacement{Part: s.main, ExpectedSHA256: mainHash, Data: data})
		result.State = "recalculation-required"
	} else {
		result.State = "caches-unchanged"
	}
	chain, err := s.calculationChain(graph)
	if err != nil {
		return CalculationEffect{}, err
	}
	if len(affected) > 0 && chain.part != "" {
		plan, err := s.pkg.PlanGraphMutation(packaging.GraphMutation{Replacements: replacements, Removals: []packaging.RelationshipRemoval{{Source: s.main, ID: chain.id}}, Deletions: []packaging.PartDeletion{{Name: chain.part, ExpectedSHA256: chain.hash}}})
		if err != nil {
			return CalculationEffect{}, err
		}
		if err = s.pkg.ApplyGraphPlan(plan); err != nil {
			return CalculationEffect{}, err
		}
		result.RemovedCalculationChain = chain.part
	} else if err = s.pkg.Replace(replacements); err != nil {
		return CalculationEffect{}, err
	}
	s.generation++
	target.consumed = true
	result.ValueChanged = true
	return result, nil
}
