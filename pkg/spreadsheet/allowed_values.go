package spreadsheet

import (
	"encoding/xml"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/formula"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ValidationValue is a stored validation vocabulary value, never a calculated or
// formatted display value. Kind is string, number, boolean or blank. Text retains
// literal/string content or the stored scalar spelling; blank has empty Text.
type ValidationValue struct {
	Kind string `json:"kind"`
	Text string `json:"text"`
}

const maxValidationValues = 100000

// ParseFloat also accepts Go hexadecimal and underscore spellings. Stored
// SpreadsheetML numbers in this API must use finite decimal/exponent syntax.
var validationNumber = regexp.MustCompile(`^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$`)

type validationRect struct{ loRow, hiRow, loCol, hiCol int }

func (r validationRect) contains(c formula.Cell) bool {
	return c.Row >= r.loRow && c.Row <= r.hiRow && c.Column >= r.loCol && c.Column <= r.hiCol
}
func validationRange(text string) (formula.Reference, error) {
	refs, err := formula.Analyze(text)
	if err != nil || len(refs) != 1 || refs[0].Start != 0 || refs[0].End != len(text) {
		return formula.Reference{}, editRefusal("unsupported_structure", "validation requires a direct static A1 cell/range")
	}
	return refs[0], nil
}
func validationBounds(r formula.Reference) validationRect {
	return validationRect{min(r.First.Row, r.Last.Row), max(r.First.Row, r.Last.Row), min(r.First.Column, r.Last.Column), max(r.First.Column, r.Last.Column)}
}
func validationLocalRange(text string) (validationRect, error) {
	r, err := validationRange(text)
	if err != nil {
		return validationRect{}, err
	}
	if r.Sheet != "" {
		return validationRect{}, editRefusal("unsupported_structure", "sheet-qualified validation target/merge unsupported")
	}
	return validationBounds(r), nil
}
func validationAttrs(e losslessxml.Element, allowed string) error {
	for _, a := range e.Attributes() {
		ok := false
		for _, name := range strings.Fields(allowed) {
			if a.Name == (xml.Name{Local: name}) {
				ok = true
				break
			}
		}
		if !ok {
			return editRefusal("unsupported_structure", "unknown validation attribute: "+a.Name.Local)
		}
	}
	return nil
}
func validationWhitespace(e losslessxml.Element) error {
	t, _ := e.Text()
	for _, c := range t {
		if c != ' ' && c != '\t' && c != '\r' && c != '\n' {
			return editRefusal("unsupported_structure", "mixed validation XML content")
		}
	}
	return nil
}
func validationLeaf(e losslessxml.Element) (string, error) {
	text, leaf := e.Text()
	if !leaf || len(e.Attributes()) != 0 {
		return "", editRefusal("unsupported_structure", "plain validation value leaf required")
	}
	return text, nil
}

type validationXML struct {
	doc      *losslessxml.Document
	root     losslessxml.Element
	children map[losslessxml.Element][]losslessxml.Element
}

func (s *EditSession) validationSheet(part string) (validationXML, error) {
	out := validationXML{}
	b, _, err := s.pkg.Part(part)
	if err != nil {
		return out, err
	}
	out.doc, err = losslessxml.Parse(b)
	if err != nil {
		return out, editRefusal("unsupported_structure", err.Error())
	}
	es := out.doc.Elements()
	out.root = es[0]
	if out.root.Name() != expanded("worksheet") {
		return out, editRefusal("unsupported_structure", "worksheet root required")
	}
	out.children = map[losslessxml.Element][]losslessxml.Element{}
	for _, e := range es {
		if e.Name().Space != packaging.NSSpreadsheetML || e.Name().Local == "extLst" {
			return out, editRefusal("unsupported_structure", "extended worksheet validation cannot be excluded")
		}
		if p, ok := e.Parent(); ok {
			out.children[p] = append(out.children[p], e)
		}
	}
	return out, nil
}
func (v validationXML) merges() ([]validationRect, error) {
	var result []validationRect
	containers := 0
	for _, e := range v.doc.Elements() {
		switch e.Name().Local {
		case "mergeCells":
			p, ok := e.Parent()
			if !ok || p != v.root {
				return nil, editRefusal("unsupported_structure", "nested merge registry")
			}
			containers++
			if containers > 1 {
				return nil, editRefusal("ambiguous_target", "multiple merge registries")
			}
			if err := validationAttrs(e, "count"); err != nil {
				return nil, err
			}
			if err := validationWhitespace(e); err != nil {
				return nil, err
			}
			children := v.children[e]
			if c, present := attributeValue(e, "count"); present {
				n, err := unsignedIndex(c)
				if err != nil || n != uint64(len(children)) {
					return nil, editRefusal("unsupported_structure", "merge count mismatch")
				}
			}
			for _, child := range children {
				if child.Name() != expanded("mergeCell") {
					return nil, editRefusal("unsupported_structure", "unknown merge entry")
				}
				if err := validationAttrs(child, "ref"); err != nil {
					return nil, err
				}
				if len(v.children[child]) != 0 {
					return nil, editRefusal("unsupported_structure", "nested merge entry")
				}
				if err := validationWhitespace(child); err != nil {
					return nil, err
				}
				r, err := validationLocalRange(attr(child, "ref"))
				if err != nil {
					return nil, err
				}
				result = append(result, r)
			}
		case "mergeCell":
			p, ok := e.Parent()
			if !ok || p.Name() != expanded("mergeCells") {
				return nil, editRefusal("unsupported_structure", "orphan merge entry")
			}
		}
	}
	return result, nil
}
func validationInterior(merges []validationRect, c formula.Cell) bool {
	for _, r := range merges {
		if r.contains(c) && (c.Row != r.loRow || c.Column != r.loCol) {
			return true
		}
	}
	return false
}

// AllowedValues inspects one worksheet cell without creating it. A nil slice
// means no list validation covers it; unproved sources return errors, never a
// partial vocabulary. This bounded subset accepts literal lists and static 1D
// ranges (including stored-cell-bounded whole axes) up to 100000 positions,
// reading plain/rich inline or shared strings, finite numbers and booleans.
// Returned values retain source order, duplicates, whitespace and blank slots.
func (s *EditSession) AllowedValues(sheet, cell string) ([]ValidationValue, error) {
	coord, err := formula.ParseCell(cell)
	if err != nil {
		return nil, editRefusal("unsupported_structure", "bounded A1 validation target required")
	}
	part := s.sheets[sheet]
	if part == "" {
		return nil, editRefusal("missing_target", "validation worksheet absent")
	}
	v, err := s.validationSheet(part)
	if err != nil {
		return nil, err
	}
	var matches []losslessxml.Element
	containers := 0
	for _, e := range v.doc.Elements() {
		switch e.Name().Local {
		case "dataValidations":
			parent, ok := e.Parent()
			if !ok || parent != v.root {
				return nil, editRefusal("unsupported_structure", "nested validation registry")
			}
			containers++
			if containers > 1 {
				return nil, editRefusal("ambiguous_target", "multiple validation registries")
			}
			if err = validationAttrs(e, "count disablePrompts xWindow yWindow"); err != nil {
				return nil, err
			}
			if err = validationWhitespace(e); err != nil {
				return nil, err
			}
			children := v.children[e]
			if count, present := attributeValue(e, "count"); present {
				n, err := unsignedIndex(count)
				if err != nil || n != uint64(len(children)) {
					return nil, editRefusal("unsupported_structure", "validation count mismatch")
				}
			}
			for _, dv := range children {
				if dv.Name() != expanded("dataValidation") {
					return nil, editRefusal("unsupported_structure", "unknown validation entry")
				}
				kind, present := attributeValue(dv, "type")
				if !present { // The schema default is none.
					continue
				}
				switch kind {
				case "none", "whole", "decimal", "date", "time", "textLength", "custom":
					continue
				case "list":
				default:
					return nil, editRefusal("unsupported_structure", "unknown validation type")
				}
				sqref := strings.Fields(attr(dv, "sqref"))
				if len(sqref) == 0 {
					return nil, editRefusal("unsupported_structure", "list validation lacks coverage")
				}
				covered := false
				for _, area := range sqref {
					r, err := validationLocalRange(area)
					if err != nil {
						return nil, err
					}
					covered = covered || r.contains(coord)
				}
				if covered {
					matches = append(matches, dv)
				}
			}
		case "dataValidation":
			parent, ok := e.Parent()
			if !ok || parent.Name() != expanded("dataValidations") {
				return nil, editRefusal("unsupported_structure", "orphan validation entry")
			}
		}
	}
	if len(matches) == 0 {
		return nil, nil
	}
	if len(matches) != 1 {
		return nil, editRefusal("ambiguous_target", "overlapping list validations")
	}
	merges, err := v.merges()
	if err != nil {
		return nil, err
	}
	if validationInterior(merges, coord) {
		return nil, editRefusal("unsupported_structure", "validation target is a merged interior")
	}
	dv := matches[0]
	if err = validationAttrs(dv, "type sqref allowBlank showDropDown showInputMessage showErrorMessage errorStyle imeMode errorTitle error promptTitle prompt operator"); err != nil {
		return nil, err
	}
	if err = validationWhitespace(dv); err != nil {
		return nil, err
	}
	children := v.children[dv]
	if len(children) != 1 || children[0].Name() != expanded("formula1") {
		return nil, editRefusal("unsupported_structure", "one list source formula required")
	}
	source, err := validationLeaf(children[0])
	if err != nil {
		return nil, err
	}
	source = strings.TrimSpace(source)
	if strings.HasPrefix(source, "=") {
		source = strings.TrimSpace(source[1:])
	}
	if source == "" {
		return nil, editRefusal("unsupported_structure", "empty validation source")
	}
	if strings.HasPrefix(source, `"`) {
		return literalValidationValues(source)
	}
	ref, err := formula.ParseRange(source)
	if err != nil {
		return nil, editRefusal("unsupported_structure", err.Error())
	}
	if ref.Sheet != "" {
		found := 0
		for name, p := range s.sheets {
			if strings.EqualFold(name, ref.Sheet) {
				part = p
				found++
			}
		}
		if found == 0 {
			return nil, editRefusal("missing_target", "validation source worksheet absent")
		}
		if found != 1 {
			return nil, editRefusal("ambiguous_target", "validation source sheet identity ambiguous")
		}
		v, err = s.validationSheet(part)
		if err != nil {
			return nil, err
		}
	}
	merges, err = v.merges()
	if err != nil {
		return nil, err
	}
	cells, err := v.storedCells()
	if err != nil {
		return nil, err
	}
	maxRow, maxCol := 1, 1
	for c := range cells {
		maxRow = max(maxRow, c[0])
		maxCol = max(maxCol, c[1])
	}
	first, last := ref.First, ref.Last
	if ref.WholeColumns {
		first.Row, last.Row = 1, maxRow
	}
	if ref.WholeRows {
		first.Column, last.Column = 1, maxCol
	}
	bounds := validationBounds(formula.Reference{First: first, Last: last})
	if bounds.loRow != bounds.hiRow && bounds.loCol != bounds.hiCol {
		return nil, editRefusal("unsupported_structure", "two-dimensional validation range")
	}
	count := max(bounds.hiRow-bounds.loRow, bounds.hiCol-bounds.loCol) + 1
	if count > maxValidationValues {
		return nil, editRefusal("resource_limit", "validation vocabulary exceeds 100000 positions")
	}
	out := make([]ValidationValue, 0, count)
	shared := &validationStringTable{}
	for i := 0; i < count; i++ {
		c := formula.Cell{Row: bounds.loRow, Column: bounds.loCol}
		if bounds.loRow == bounds.hiRow {
			c.Column += i
		} else {
			c.Row += i
		}
		if validationInterior(merges, c) {
			return nil, editRefusal("unsupported_structure", "validation source includes merged interior")
		}
		e, ok := cells[[2]int{c.Row, c.Column}]
		value := ValidationValue{Kind: "blank"}
		if ok {
			value, err = v.storedValue(s, e, shared)
			if err != nil {
				return nil, err
			}
		}
		out = append(out, value)
	}
	return out, nil
}
func literalValidationValues(source string) ([]ValidationValue, error) {
	if len(source) < 2 || source[len(source)-1] != '"' {
		return nil, editRefusal("unsupported_structure", "unterminated validation literal")
	}
	inside := source[1 : len(source)-1]
	if strings.Count(inside, ",") >= maxValidationValues {
		return nil, editRefusal("resource_limit", "validation vocabulary exceeds 100000 positions")
	}
	items := []string{""}
	var b strings.Builder
	for i := 0; i < len(inside); i++ {
		switch inside[i] {
		case ',':
			items[len(items)-1] = b.String()
			b.Reset()
			items = append(items, "")
		case '"':
			if i+1 >= len(inside) || inside[i+1] != '"' {
				return nil, editRefusal("unsupported_structure", "invalid quote in validation literal")
			}
			b.WriteByte('"')
			i++
		default:
			b.WriteByte(inside[i])
		}
	}
	items[len(items)-1] = b.String()
	out := make([]ValidationValue, len(items))
	for i, text := range items {
		out[i] = ValidationValue{Kind: "string", Text: text}
	}
	return out, nil
}
func (v validationXML) storedCells() (map[[2]int]losslessxml.Element, error) {
	cells := map[[2]int]losslessxml.Element{}
	sheetData := 0
	rowIDs := map[uint64]bool{}
	for _, e := range v.doc.Elements() {
		switch e.Name().Local {
		case "sheetData":
			p, ok := e.Parent()
			if !ok || p != v.root {
				return nil, editRefusal("unsupported_structure", "nested sheetData")
			}
			sheetData++
		case "row":
			p, ok := e.Parent()
			if !ok || p.Name() != expanded("sheetData") {
				return nil, editRefusal("unsupported_structure", "orphan row")
			}
			id, err := unsignedIndex(attr(e, "r"))
			if err != nil || id < 1 || id > 1048576 || rowIDs[id] {
				return nil, editRefusal("ambiguous_target", "invalid/duplicate row identity")
			}
			rowIDs[id] = true
		case "c":
			ref := attr(e, "r")
			c, err := formula.ParseCell(ref)
			if err != nil || !strictCell.MatchString(ref) {
				return nil, editRefusal("unsupported_structure", "canonical stored cell reference required")
			}
			p, ok := e.Parent()
			if !ok || p.Name() != expanded("row") {
				return nil, editRefusal("unsupported_structure", "cell owner not row")
			}
			row, err := unsignedIndex(attr(p, "r"))
			if err != nil || row != uint64(c.Row) {
				return nil, editRefusal("unsupported_structure", "row/cell mismatch")
			}
			key := [2]int{c.Row, c.Column}
			if _, exists := cells[key]; exists {
				return nil, editRefusal("ambiguous_target", "duplicate source cell")
			}
			cells[key] = e
		}
	}
	if sheetData != 1 {
		return nil, editRefusal("unsupported_structure", "one sheetData required")
	}
	return cells, nil
}
func (v validationXML) storedValue(session *EditSession, cell losslessxml.Element, shared *validationStringTable) (ValidationValue, error) {
	fail := func(detail string) (ValidationValue, error) {
		return ValidationValue{}, editRefusal("unsupported_structure", detail)
	}
	if err := validationAttrs(cell, "r s t"); err != nil {
		return ValidationValue{}, err
	}
	if err := validationWhitespace(cell); err != nil {
		return ValidationValue{}, err
	}
	children := v.children[cell]
	for _, e := range children {
		if e.Name() == expanded("f") {
			return fail("formula validation source cannot use a cached result")
		}
	}
	kind := attr(cell, "t")
	if kind == "e" {
		return fail("error validation source")
	}
	if len(children) == 0 {
		if kind == "" || kind == "n" {
			return ValidationValue{Kind: "blank"}, nil
		}
		return fail("typed source value absent")
	}
	if len(children) != 1 {
		return fail("ambiguous validation source value")
	}
	value := children[0]
	if kind == "inlineStr" {
		if value.Name() != expanded("is") || len(value.Attributes()) != 0 {
			return fail("inline string structure unsupported")
		}
		text, err := v.stringValue(value)
		if err != nil {
			return ValidationValue{}, err
		}
		return ValidationValue{Kind: "string", Text: text}, nil
	}
	if value.Name() != expanded("v") {
		return fail("value leaf required")
	}
	text, err := validationLeaf(value)
	if err != nil {
		return ValidationValue{}, err
	}
	switch kind {
	case "s":
		index, err := unsignedIndex(text)
		if err != nil {
			return fail("malformed shared string index")
		}
		if !shared.loaded {
			shared.values, err = session.validationSharedStrings()
			if err != nil {
				return ValidationValue{}, err
			}
			shared.loaded = true
		}
		if index >= uint64(len(shared.values)) {
			return fail("shared string index out of range")
		}
		return ValidationValue{Kind: "string", Text: shared.values[index]}, nil
	case "", "n":
		if !validationNumber.MatchString(text) {
			return fail("non-decimal stored numeric value")
		}
		n, err := strconv.ParseFloat(text, 64)
		if err != nil || math.IsNaN(n) || math.IsInf(n, 0) {
			return fail("nonfinite/malformed stored numeric value")
		}
		return ValidationValue{Kind: "number", Text: text}, nil
	case "b":
		if text != "0" && text != "1" {
			return fail("invalid stored boolean")
		}
		return ValidationValue{Kind: "boolean", Text: text}, nil
	default:
		return fail("unproved stored validation value type")
	}
}
