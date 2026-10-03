package spreadsheet

import (
	"encoding/xml"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// EditSession supports existing numeric cells in a conservative formula-free
// workbook subset. It does not guess dependencies or deliver stale formula caches.
// Sessions are single-owner; legacy Workbook interfaces are unchanged.
type EditSession struct {
	pkg        *packaging.Preserved
	main       string
	sheets     map[string]string
	generation uint64
}
type NumberTarget struct {
	session          *EditSession
	generation       uint64
	doc              *losslessxml.Document
	element          losslessxml.Element
	part, hash, text string
	consumed         bool
}

// CurrentNumber observes a held target without creating a new selection or
// bypassing freshness checks. It does not interpret formula cached values.
func (s *EditSession) CurrentNumber(target *NumberTarget) (float64, error) {
	if target == nil || target.session != s || target.generation != s.generation || target.consumed {
		return 0, editRefusal("stale_target", "stale, foreign or consumed numeric target")
	}
	_, hash, err := s.pkg.Part(target.part)
	if err != nil {
		return 0, err
	}
	if hash != target.hash {
		return 0, editRefusal("stale_target", "part fingerprint changed")
	}
	return strconv.ParseFloat(target.text, 64)
}

func editRefusal(kind, detail string) error {
	return &packaging.Refusal{Kind: kind, Operation: "spreadsheet_edit", Detail: detail}
}
func expanded(local string) xml.Name { return xml.Name{Space: packaging.NSSpreadsheetML, Local: local} }
func attr(e losslessxml.Element, local string) string {
	for _, a := range e.Attributes() {
		if a.Name == (xml.Name{Local: local}) {
			return a.Value
		}
	}
	return ""
}
func OpenEditing(source []byte, limits packaging.Limits) (*EditSession, error) {
	p, err := packaging.OpenPreserved(source, limits)
	if err != nil {
		return nil, err
	}
	g, err := p.Graph()
	if err != nil {
		return nil, err
	}
	main := ""
	for _, e := range g.Edges {
		if e.Source == "" && e.Type == packaging.RelTypeOfficeDocument && !e.External {
			if main != "" {
				return nil, editRefusal("ambiguous_target", "multiple main parts")
			}
			main = e.ResolvedPart
		}
	}
	if main == "" {
		return nil, editRefusal("missing_target", "workbook main part missing")
	}
	types := map[string]string{}
	for _, p := range g.Parts {
		types[p.Name] = p.ContentType
	}
	if types[main] != packaging.ContentTypeWorkbook {
		return nil, editRefusal("unsupported_structure", "only ordinary transitional XLSX supported")
	}
	data, _, err := p.Part(main)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	if d.Elements()[0].Name() != expanded("workbook") {
		return nil, editRefusal("unsupported_structure", "workbook root required")
	}
	rels := map[string]string{}
	for _, e := range g.Edges {
		if e.Source == main && e.Type == packaging.RelTypeWorksheet {
			if e.External {
				return nil, editRefusal("relationship_policy", "external sheet")
			}
			rels[e.ID] = e.ResolvedPart
		}
	}
	sheets := map[string]string{}
	parts := map[string]bool{}
	for _, e := range d.Elements() {
		if e.Name() != expanded("sheet") {
			continue
		}
		parent, ok := e.Parent()
		if !ok || parent.Name() != expanded("sheets") {
			return nil, editRefusal("unsupported_structure", "unknown sheet inventory")
		}
		sheet := attr(e, "name")
		rid := ""
		for _, a := range e.Attributes() {
			if a.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}) {
				rid = a.Value
			}
		}
		part := rels[rid]
		if sheet == "" || part == "" || sheets[sheet] != "" || parts[part] || types[part] != packaging.ContentTypeWorksheet {
			return nil, &packaging.Refusal{Kind: "xlsx-workbook-invalid", Operation: "workbook_open", Part: main, Detail: "ambiguous or wrong-URI sheet relationship"}
		}
		sheets[sheet] = part
		parts[part] = true
	}
	return &EditSession{pkg: p, main: main, sheets: sheets}, nil
}

var strictCell = regexp.MustCompile(`^[A-Z]{1,3}[1-9][0-9]*$`)

func (s *EditSession) FindNumber(sheet, cell string) (*NumberTarget, error) {
	ref, err := utils.ParseCellRef(cell)
	if err != nil || !strictCell.MatchString(cell) || ref.Row > 1048576 || ref.Col > 16384 {
		return nil, editRefusal("unsupported_structure", "expected exact bounded A1 address")
	}
	part := s.sheets[sheet]
	if part == "" {
		return nil, editRefusal("missing_target", "sheet absent")
	}
	b, hash, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(b)
	if err != nil {
		return nil, err
	}
	if d.Elements()[0].Name() != expanded("worksheet") {
		return nil, editRefusal("unsupported_structure", "worksheet root required")
	}
	var cells []losslessxml.Element
	for _, e := range d.Elements() {
		if e.Name() == expanded("c") && attr(e, "r") == cell {
			cells = append(cells, e)
		}
	}
	if len(cells) == 0 {
		return nil, editRefusal("missing_target", "existing numeric cell absent")
	}
	if len(cells) != 1 {
		return nil, editRefusal("ambiguous_target", "duplicate cell address")
	}
	c := cells[0]
	if kind := attr(c, "t"); kind != "" && kind != "n" {
		return nil, editRefusal("unsupported_structure", "cell is not numeric")
	}
	row, ok := c.Parent()
	if !ok || row.Name() != expanded("row") {
		return nil, editRefusal("unsupported_structure", "cell owner is not row")
	}
	data, ok := row.Parent()
	if !ok || data.Name() != expanded("sheetData") {
		return nil, editRefusal("unsupported_structure", "row owner is not sheetData")
	}
	root, ok := data.Parent()
	if !ok || root != d.Elements()[0] {
		return nil, editRefusal("unsupported_structure", "unknown sheetData owner")
	}
	var values []losslessxml.Element
	for _, e := range d.Elements() {
		p, ok := e.Parent()
		if ok && p == c {
			if e.Name() == expanded("f") {
				return nil, editRefusal("unsupported_structure", "cell contains formula or extension")
			}
			if e.Name() != expanded("v") {
				return nil, editRefusal("unsupported_structure", "cell contains formula or extension")
			}
			values = append(values, e)
		}
	}
	if len(values) != 1 {
		return nil, editRefusal("unsupported_structure", "numeric cell needs one value leaf")
	}
	text, leaf := values[0].Text()
	number, err := strconv.ParseFloat(text, 64)
	if !leaf || err != nil || math.IsNaN(number) || math.IsInf(number, 0) {
		return nil, editRefusal("unsupported_structure", "invalid numeric value")
	}
	return &NumberTarget{session: s, generation: s.generation, doc: d, element: values[0], part: part, hash: hash, text: text}, nil
}

// SetNumberAt is an opt-in direct numeric write. A formula cell is selected
// by its literal bounded address before its attributed topology is classified.
// The legacy FindNumber contract remains formula-free; no target is forged.
func (s *EditSession) SetNumberAt(sheet, address string, value float64) error {
	ref, err := utils.ParseCellRef(address)
	if err != nil || !strictCell.MatchString(address) || ref.Row > 1048576 || ref.Col > 16384 {
		return editRefusal("unsupported_structure", "expected exact bounded A1 address")
	}
	part := s.sheets[sheet]
	if part == "" {
		return editRefusal("missing_target", "sheet absent")
	}
	data, _, err := s.pkg.Part(part)
	if err != nil {
		return err
	}
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return err
	}
	if doc.Elements()[0].Name() != expanded("worksheet") {
		return editRefusal("unsupported_structure", "worksheet root required")
	}
	var cell losslessxml.Element
	count := 0
	for _, e := range doc.Elements() {
		if e.Name() == expanded("c") && attr(e, "r") == address {
			cell = e
			count++
		}
	}
	if count != 1 {
		if count == 0 {
			return editRefusal("missing_target", "cell absent")
		}
		return editRefusal("ambiguous_target", "duplicate cell address")
	}
	row, ok := cell.Parent()
	if !ok || row.Name() != expanded("row") {
		return editRefusal("unsupported_structure", "cell owner is not row")
	}
	sheetData, ok := row.Parent()
	if !ok || sheetData.Name() != expanded("sheetData") {
		return editRefusal("unsupported_structure", "row owner is not sheetData")
	}
	root, ok := sheetData.Parent()
	if !ok || root != doc.Elements()[0] {
		return editRefusal("unsupported_structure", "unknown sheetData owner")
	}
	for _, e := range doc.Elements() {
		if parent, ok := e.Parent(); ok && parent == cell && e.Name() == expanded("f") {
			return attributedFormulaRefusal(e)
		}
	}
	target, err := s.FindNumber(sheet, address)
	if err != nil {
		return err
	}
	return s.SetNumber(target, value)
}

// The formula topology is read directly from the selected cell, not inferred
// from an error string or from formula parsing. Refusal leaves the session and
// all unrelated held targets usable.
func attributedFormulaRefusal(formula losslessxml.Element) error {
	switch attr(formula, "t") {
	case "shared":
		return editRefusal("xlsx-shared-formula-edit-unsupported", "direct shared-formula overwrite")
	case "array":
		return editRefusal("xlsx-array-formula-edit-unsupported", "direct array-formula overwrite")
	default:
		return editRefusal("unsupported_structure", "cell contains formula or extension")
	}
}

func (s *EditSession) guard() error { return s.guardMode(false) }
func (s *EditSession) guardMode(allowStaticFormulas bool) error {
	return s.guardModeWithOpaque(allowStaticFormulas, false)
}
func (s *EditSession) guardModeWithOpaque(allowStaticFormulas, sealedOpaque bool) error {
	if err := s.ValidateStyles(); err != nil {
		return err
	}
	g, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	allowedRels := map[string]bool{packaging.RelTypeOfficeDocument: true, packaging.RelTypeWorksheet: true, packaging.RelTypeStyles: true, packaging.RelTypeSharedStrings: true, packaging.RelTypeTheme: true}
	if allowStaticFormulas {
		if _, err := s.calculationChain(g); err != nil {
			return err
		}
		allowedRels[packaging.RelTypeCalcChain] = true
	}
	for _, e := range g.Edges {
		if e.Source != "" && !allowedRels[e.Type] {
			return editRefusal("unsupported_structure", "unproved workbook relationship dependency")
		}
	}
	for _, p := range g.Parts {
		if sealedOpaque && (p.Name == "xl/charts/cache-boundary.xml" || p.Name == "xl/externalLinks/cache-boundary.xml") {
			continue
		}
		if strings.HasPrefix(p.Name, "xl/") && (strings.Contains(p.Name, "/charts/") || strings.Contains(p.Name, "/pivot") || strings.Contains(p.Name, "/external") || strings.Contains(p.Name, "/tables/") || strings.Contains(p.Name, "/connections") || (!allowStaticFormulas && strings.EqualFold(p.Name, "xl/calcChain.xml"))) {
			return editRefusal("unsupported_structure", "unproved chart/pivot/table/external dependency")
		}
		isSheet := p.ContentType == packaging.ContentTypeWorksheet
		if p.Name != s.main && !isSheet {
			continue
		}
		b, _, err := s.pkg.Part(p.Name)
		if err != nil {
			return err
		}
		d, err := losslessxml.Parse(b)
		if err != nil {
			return err
		}
		allowed := map[string]bool{}
		for _, local := range strings.Fields("workbook fileVersion workbookPr bookViews workbookView sheets sheet calcPr worksheet sheetPr tabColor outlinePr pageSetUpPr dimension sheetViews sheetView pane selection sheetFormatPr cols col sheetData row c v is t r rPr phoneticPr pageMargins pageSetup printOptions headerFooter oddHeader oddFooter evenHeader evenFooter firstHeader firstFooter sheetCalcPr") {
			allowed[local] = true
		}
		if allowStaticFormulas {
			allowed["f"] = true
		}
		for _, e := range d.Elements() {
			n := e.Name()
			for _, a := range e.Attributes() {
				if a.Name.Space != "" && !(n.Local == "sheet" && a.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"})) && !(a.Name == (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"})) {
					return editRefusal("unsupported_structure", "unknown namespaced attribute")
				}
			}
			if n.Space != packaging.NSSpreadsheetML {
				return editRefusal("unsupported_structure", "unknown workbook/worksheet extension")
			}
			if (n.Local == "workbookProtection" || n.Local == "definedNames") && len(e.Attributes()) == 0 {
				text, leaf := e.Text()
				if leaf && strings.TrimSpace(text) == "" {
					continue
				}
			}
			if n.Local == "f" && !allowStaticFormulas {
				return editRefusal("unsupported_structure", "formulas require explicit cache invalidation")
			}
			switch n.Local {
			case "definedName", "definedNames", "dataValidation", "dataValidations", "conditionalFormatting", "mergeCells", "mergeCell", "tableParts", "drawing", "legacyDrawing", "extLst", "externalReferences", "pivotCaches", "calcChain", "oleObjects", "sheetProtection", "workbookProtection":
				kind := "unsupported_structure"
				if n.Local == "sheetProtection" || n.Local == "workbookProtection" {
					kind = "protected_operation"
				}
				return editRefusal(kind, "dependent/protected structure requires reference-aware editor: "+n.Local)
			}
			if !allowed[n.Local] {
				return editRefusal("unsupported_structure", "unclassified workbook/worksheet element: "+n.Local)
			}
		}
	}
	return nil
}
func (s *EditSession) SetNumber(target *NumberTarget, value float64) error {
	if target == nil || target.session != s || target.generation != s.generation || target.consumed {
		return editRefusal("stale_target", "stale, foreign or consumed target")
	}
	if math.IsNaN(value) || math.IsInf(value, 0) {
		return editRefusal("unsupported_structure", "finite number required")
	}
	_, hash, err := s.pkg.Part(target.part)
	if err != nil {
		return err
	}
	if hash != target.hash {
		return editRefusal("stale_target", "part fingerprint changed")
	}
	if err = s.guard(); err != nil {
		return err
	}
	old, _ := strconv.ParseFloat(target.text, 64)
	if old == value {
		return nil
	}
	text := strconv.FormatFloat(value, 'g', -1, 64)
	b, err := target.doc.ReplaceText([]losslessxml.TextEdit{{Target: target.element, Text: text}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: target.part, ExpectedSHA256: target.hash, Data: b}}); err != nil {
		return err
	}
	s.generation++
	target.consumed = true
	return nil
}
func (s *EditSession) SaveAs(path string) (packaging.Receipt, error) { return s.pkg.SaveAs(path) }
