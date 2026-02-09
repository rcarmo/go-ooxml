// Package spreadsheet provides a high-level API for working with Excel workbooks.
package spreadsheet

import (
	"fmt"
	"io"
	"path"
	"path/filepath"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// workbookImpl represents an Excel workbook.
type workbookImpl struct {
	pkg           *packaging.Package
	workbook      *sml.Workbook
	sheets        []*worksheetImpl
	sharedStrings *SharedStrings
	path          string
	nextSheetID   int
	nextTableID   int
	comments      map[string]*SheetComments
	styles        *stylesImpl
	themeParts    map[string][]byte
	extraParts    map[string]*packaging.Part
}

// New creates a new empty workbook with one sheet.
func New() (Workbook, error) {
	w := &workbookImpl{
		pkg: packaging.New(),
		workbook: &sml.Workbook{
			Sheets: &sml.Sheets{},
			BookViews: &sml.BookViews{
				WorkbookView: []*sml.WorkbookView{{
					XWindow:      0,
					YWindow:      0,
					WindowWidth:  28800,
					WindowHeight: 12300,
				}},
			},
		},
		sheets:        make([]*worksheetImpl, 0),
		sharedStrings: newSharedStrings(),
		nextSheetID:   1,
		nextTableID:   1,
		comments:      make(map[string]*SheetComments),
		styles:        newStyles(nil),
		themeParts:    make(map[string][]byte),
		extraParts:    make(map[string]*packaging.Part),
	}

	if err := w.initPackage(); err != nil {
		return nil, err
	}

	// Add default sheet
	w.AddSheet("Sheet1")

	return w, nil
}

// Open opens an existing workbook from a file path.
func Open(path string) (Workbook, error) {
	pkg, err := packaging.Open(path)
	if err != nil {
		return nil, err
	}

	w, err := openFromPackage(pkg)
	if err != nil {
		return nil, err
	}
	w.path = path
	return w, nil
}

// OpenReader opens a workbook from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (Workbook, error) {
	pkg, err := packaging.OpenReader(r, size)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

func openFromPackage(pkg *packaging.Package) (*workbookImpl, error) {
	w := &workbookImpl{
		pkg:           pkg,
		sheets:        make([]*worksheetImpl, 0),
		sharedStrings: newSharedStrings(),
		nextSheetID:   1,
		nextTableID:   1,
		comments:      make(map[string]*SheetComments),
		styles:        newStyles(nil),
		themeParts:    make(map[string][]byte),
		extraParts:    make(map[string]*packaging.Part),
	}

	// Parse workbook.xml
	if err := w.parseWorkbook(); err != nil {
		return nil, err
	}

	// Parse shared strings
	w.parseSharedStrings()

	// Parse styles
	w.parseStyles()

	// Parse worksheets
	if err := w.parseSheets(); err != nil {
		return nil, err
	}

	// Parse worksheet comments
	w.parseComments()

	w.captureAdvancedParts()

	return w, nil
}

// Save saves the workbook to its original path.
func (w *workbookImpl) Save() error {
	if w.path == "" {
		return utils.ErrPathNotSet
	}
	return w.SaveAs(w.path)
}

// SaveAs saves the workbook to a new path.
func (w *workbookImpl) SaveAs(path string) error {
	if err := w.updatePackage(); err != nil {
		return err
	}
	return w.pkg.SaveAs(path)
}

// Close closes the workbook and releases resources.
func (w *workbookImpl) Close() error {
	return w.pkg.Close()
}

// CoreProperties returns the workbook core properties.
func (w *workbookImpl) CoreProperties() (*common.CoreProperties, error) {
	return w.pkg.CoreProperties()
}

// SetCoreProperties sets the workbook core properties.
func (w *workbookImpl) SetCoreProperties(props *common.CoreProperties) error {
	return w.pkg.SetCoreProperties(props)
}

// Styles returns the workbook styles manager.
func (w *workbookImpl) Styles() Styles {
	if w.styles == nil {
		w.styles = newStyles(nil)
	}
	return w.styles
}

// =============================================================================
// Sheet access
// =============================================================================

// Sheets returns all worksheets in the workbook.
func (w *workbookImpl) Sheets() []Worksheet {
	sheets := make([]Worksheet, len(w.sheets))
	for i, sheet := range w.sheets {
		sheets[i] = sheet
	}
	return sheets
}

// SheetsRaw returns all worksheets in the workbook as concrete types.
func (w *workbookImpl) SheetsRaw() []*worksheetImpl {
	return w.sheets
}

// SheetRaw returns a worksheet by name or index as the concrete type.
func (w *workbookImpl) SheetRaw(nameOrIndex interface{}) (*worksheetImpl, error) {
	switch id := nameOrIndex.(type) {
	case int:
		if id < 0 || id >= len(w.sheets) {
			return nil, ErrSheetNotFound
		}
		return w.sheets[id], nil
	case string:
		for _, sheet := range w.sheets {
			if sheet.Name() == id {
				return sheet, nil
			}
		}
		return nil, ErrSheetNotFound
	default:
		return nil, ErrSheetNotFound
	}
}

// Sheet returns a worksheet by name or index.
func (w *workbookImpl) Sheet(nameOrIndex interface{}) (Worksheet, error) {
	switch id := nameOrIndex.(type) {
	case int:
		if id < 0 || id >= len(w.sheets) {
			return nil, ErrSheetNotFound
		}
		return w.sheets[id], nil
	case string:
		for _, sheet := range w.sheets {
			if sheet.Name() == id {
				return sheet, nil
			}
		}
		return nil, ErrSheetNotFound
	default:
		return nil, ErrSheetNotFound
	}
}

// SheetCount returns the number of worksheets.
func (w *workbookImpl) SheetCount() int {
	return len(w.sheets)
}

// SharedStrings returns the shared strings table.
func (w *workbookImpl) SharedStrings() *SharedStrings {
	return w.sharedStrings
}

// NamedRanges returns all defined names in the workbook.
func (w *workbookImpl) NamedRanges() []NamedRange {
	if w.workbook == nil || w.workbook.DefinedNames == nil {
		return nil
	}
	ranges := make([]NamedRange, len(w.workbook.DefinedNames.DefinedName))
	for i, def := range w.workbook.DefinedNames.DefinedName {
		ranges[i] = &namedRangeImpl{workbook: w, definedName: def}
	}
	return ranges
}

// AddNamedRange adds a new named range to the workbook.
func (w *workbookImpl) AddNamedRange(name, refersTo string) NamedRange {
	if w.workbook.DefinedNames == nil {
		w.workbook.DefinedNames = &sml.DefinedNames{}
	}
	if name == "" {
		name = w.nextNamedRangeName()
	}
	def := &sml.DefinedName{
		Name:  name,
		Value: refersTo,
	}
	w.workbook.DefinedNames.DefinedName = append(w.workbook.DefinedNames.DefinedName, def)
	return &namedRangeImpl{workbook: w, definedName: def}
}

func (w *workbookImpl) nextNamedRangeName() string {
	index := 1
	if w.workbook != nil && w.workbook.DefinedNames != nil {
		index = len(w.workbook.DefinedNames.DefinedName) + 1
	}
	for {
		name := fmt.Sprintf("NamedRange%d", index)
		if !w.namedRangeExists(name) {
			return name
		}
		index++
	}
}

func (w *workbookImpl) namedRangeExists(name string) bool {
	if w.workbook == nil || w.workbook.DefinedNames == nil {
		return false
	}
	for _, def := range w.workbook.DefinedNames.DefinedName {
		if def.Name == name {
			return true
		}
	}
	return false
}

// Tables returns all tables across all worksheets.
func (w *workbookImpl) Tables() []Table {
	var tables []Table
	for _, sheet := range w.sheets {
		for _, table := range sheet.Tables() {
			tables = append(tables, table)
		}
	}
	return tables
}

// Table returns a table by name.
func (w *workbookImpl) Table(name string) (Table, error) {
	for _, table := range w.Tables() {
		if table.Name() == name {
			return table, nil
		}
	}
	return nil, ErrTableNotFound
}

// AddSheet adds a new worksheet with the given name.
func (w *workbookImpl) AddSheet(name string) Worksheet {
	relID := fmt.Sprintf("rId%d", len(w.sheets)+1)

	worksheet := &sml.Worksheet{
		SheetViews: &sml.SheetViews{
			SheetView: []*sml.SheetView{{
				WorkbookViewID: 0,
			}},
		},
		SheetFormatPr: &sml.SheetFormatPr{
			DefaultRowHeight: 15,
		},
		SheetData: &sml.SheetData{},
	}

		sheet := &worksheetImpl{
			workbook:  w,
			worksheet: worksheet,
			name:      name,
			sheetID:   w.nextSheetID,
			relID:     relID,
			index:     len(w.sheets),
		}
	w.nextSheetID++

	w.sheets = append(w.sheets, sheet)

	// Add to workbook
	w.workbook.Sheets.Sheet = append(w.workbook.Sheets.Sheet, &sml.Sheet{
		Name:    name,
		SheetID: sheet.sheetID,
		ID:      relID,
	})

	return sheet
}

// DeleteSheet removes a worksheet by name or index.
func (w *workbookImpl) DeleteSheet(nameOrIndex interface{}) error {
	var index int

	switch id := nameOrIndex.(type) {
	case int:
		if id < 0 || id >= len(w.sheets) {
			return ErrSheetNotFound
		}
		index = id
	case string:
		found := false
		for i, sheet := range w.sheets {
			if sheet.Name() == id {
				index = i
				found = true
				break
			}
		}
		if !found {
			return ErrSheetNotFound
		}
	default:
		return ErrSheetNotFound
	}

	// Don't allow deleting the last sheet
	if len(w.sheets) == 1 {
		return utils.ErrCannotDeleteLastSheet
	}

	// Remove from sheets slice
	w.sheets = append(w.sheets[:index], w.sheets[index+1:]...)

	// Update indices
	for i := index; i < len(w.sheets); i++ {
		w.sheets[i].index = i
	}

	// Remove from workbook
	if index < len(w.workbook.Sheets.Sheet) {
		w.workbook.Sheets.Sheet = append(
			w.workbook.Sheets.Sheet[:index],
			w.workbook.Sheets.Sheet[index+1:]...,
		)
	}

	return nil
}

// =============================================================================
// Internal methods
// =============================================================================

func (w *workbookImpl) initPackage() error {
	// Add main relationship
	w.pkg.AddRelationship("/", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
	return nil
}

func (w *workbookImpl) parseWorkbook() error {
	part, err := w.pkg.GetPart(packaging.ExcelWorkbookPath)
	if err != nil {
		return err
	}
	data, err := part.Content()
	if err != nil {
		return err
	}

	w.workbook = &sml.Workbook{}
	return utils.UnmarshalXML(data, w.workbook)
}

func (w *workbookImpl) parseSharedStrings() {
	part, err := w.pkg.GetPart(packaging.ExcelSharedStringsPath)
	if err != nil {
		return // Shared strings are optional
	}
	data, err := part.Content()
	if err != nil {
		return
	}

	if err := w.sharedStrings.parse(data); err != nil {
		return
	}
}

func (w *workbookImpl) parseStyles() {
	part, err := w.pkg.GetPart(packaging.ExcelStylesPath)
	if err != nil {
		return
	}
	data, err := part.Content()
	if err != nil {
		return
	}
	stylesXML := &sml.StyleSheet{}
	if err := utils.UnmarshalXML(data, stylesXML); err != nil {
		return
	}
	w.styles = newStyles(stylesXML)
}

func (w *workbookImpl) parseSheets() error {
	if w.workbook.Sheets == nil {
		return nil
	}

	// Get relationships
	rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
	if rels == nil {
		return nil
	}

	for i, sheetRef := range w.workbook.Sheets.Sheet {
		// Find the relationship
		var sheetPath string
		for _, rel := range rels.Relationships {
			if rel.ID == sheetRef.ID {
				sheetPath = packaging.ResolveRelationshipTarget(packaging.ExcelWorkbookPath, rel.Target)
				if !strings.HasPrefix(sheetPath, "xl/") {
					sheetPath = "xl/" + strings.TrimPrefix(sheetPath, "/")
				}
				break
			}
		}

		if sheetPath == "" {
			continue
		}

		// Parse worksheet
		part, err := w.pkg.GetPart(sheetPath)
		if err != nil {
			continue
		}
		data, err := part.Content()
		if err != nil {
			continue
		}

		worksheet := &sml.Worksheet{}
		if err := utils.UnmarshalXML(data, worksheet); err != nil {
			continue
		}

		sheet := &worksheetImpl{
			workbook:  w,
			worksheet: worksheet,
			name:      sheetRef.Name,
			sheetID:   sheetRef.SheetID,
			relID:     sheetRef.ID,
			index:     i,
			path:      sheetPath,
			comments:  w.commentsForSheet(sheetPath),
		}

		if err := w.parseTables(sheet); err != nil {
			return err
		}

		w.sheets = append(w.sheets, sheet)

		if sheetRef.SheetID >= w.nextSheetID {
			w.nextSheetID = sheetRef.SheetID + 1
		}
	}

	return nil
}

func (w *workbookImpl) parseComments() {
	for _, sheet := range w.sheets {
		sheet.comments = w.commentsForSheet(sheet.path)
	}
}

func (w *workbookImpl) commentsForSheet(sheetPath string) *SheetComments {
	if sheetPath == "" {
		return nil
	}
	if comments, ok := w.comments[sheetPath]; ok {
		return comments
	}
	rels := w.pkg.GetRelationships(sheetPath)
	if rels == nil {
		return nil
	}
	commentRel := rels.FirstByType(packaging.RelTypeComments)
	if commentRel == nil {
		return nil
	}
	vmlRel := rels.FirstByType(packaging.RelTypeVML)
	commentPath := packaging.ResolveRelationshipTarget(sheetPath, commentRel.Target)
	part, err := w.pkg.GetPart(commentPath)
	if err != nil {
		return nil
	}
	data, err := part.Content()
	if err != nil {
		return nil
	}
	commentsXML := &sml.Comments{}
	if err := utils.UnmarshalXML(data, commentsXML); err != nil {
		return nil
	}
	comments := newSheetComments(commentPath, commentRel.ID, commentsXML)
	if vmlRel != nil {
		comments.vmlPath = packaging.ResolveRelationshipTarget(sheetPath, vmlRel.Target)
		comments.vmlRelID = vmlRel.ID
	}
	w.comments[sheetPath] = comments
	return comments
}

func (w *workbookImpl) parseTables(sheet *worksheetImpl) error {
	if sheet.worksheet.TableParts == nil || len(sheet.worksheet.TableParts.TablePart) == 0 {
		return nil
	}

	sheetRels := w.pkg.GetRelationships(sheet.path)
	if sheetRels == nil {
		return nil
	}

	for _, tablePart := range sheet.worksheet.TableParts.TablePart {
		rel := sheetRels.ByID(tablePart.ID)
		if rel == nil {
			continue
		}
		tablePath := packaging.ResolveRelationshipTarget(sheet.path, rel.Target)
		part, err := w.pkg.GetPart(tablePath)
		if err != nil {
			continue
		}
		data, err := part.Content()
		if err != nil {
			continue
		}
		tableXML := &sml.Table{}
		if err := utils.UnmarshalXML(data, tableXML); err != nil {
			continue
		}
		table := &tableImpl{
			worksheet: sheet,
			table:     tableXML,
			relID:     tablePart.ID,
			path:      tablePath,
		}
		sheet.tables = append(sheet.tables, table)
		if tableXML.ID >= w.nextTableID {
			w.nextTableID = tableXML.ID + 1
		}
	}
	return nil
}

func (w *workbookImpl) captureAdvancedParts() {
	if w.pkg == nil {
		return
	}
	workbookRels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
	for _, rel := range workbookRels.ByType(packaging.RelTypeTheme) {
		target := packaging.ResolveRelationshipTarget(packaging.ExcelWorkbookPath, rel.Target)
		part, err := w.pkg.GetPart(target)
		if err != nil {
			continue
		}
		if data, err := part.Content(); err == nil {
			w.themeParts[target] = data
		}
	}
	for _, sheet := range w.sheets {
		rels := w.pkg.GetRelationships(sheet.path)
		for _, rel := range rels.Relationships {
			switch rel.Type {
			case packaging.RelTypeDrawing, packaging.RelTypeVML:
			default:
				continue
			}
			target := packaging.ResolveRelationshipTarget(sheet.path, rel.Target)
			part, err := w.pkg.GetPart(target)
			if err != nil {
				continue
			}
			w.extraParts[target] = part
			w.captureRelatedParts(target, 2)
		}
	}
}

func (w *workbookImpl) captureRelatedParts(sourcePath string, depth int) {
	if depth <= 0 || w.pkg == nil {
		return
	}
	rels := w.pkg.GetRelationships(sourcePath)
	for _, rel := range rels.Relationships {
		switch rel.Type {
		case packaging.RelTypeChart, packaging.RelTypeChartStyle, packaging.RelTypeChartColorStyle,
			packaging.RelTypeDiagramData, packaging.RelTypeDiagramLayout,
			packaging.RelTypeDiagramColors, packaging.RelTypeDiagramStyle,
			packaging.RelTypeImage, packaging.RelTypeAudio, packaging.RelTypeVideo, packaging.RelTypeMedia:
		default:
			continue
		}
		target := packaging.ResolveRelationshipTarget(sourcePath, rel.Target)
		if _, ok := w.extraParts[target]; ok {
			continue
		}
		part, err := w.pkg.GetPart(target)
		if err != nil {
			continue
		}
		w.extraParts[target] = part
		w.captureRelatedParts(target, depth-1)
	}
}

func (w *workbookImpl) updatePackage() error {
	// Save workbook.xml
	data, err := utils.MarshalXMLWithHeader(w.workbook)
	if err != nil {
		return err
	}
	if _, err := w.pkg.AddPart(packaging.ExcelWorkbookPath, packaging.ContentTypeWorkbook, data); err != nil {
		return err
	}

		// Save each worksheet
		for i, sheet := range w.sheets {
			sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)

			if err := w.updateTables(sheet, i+1); err != nil {
				return err
			}

			if err := w.updateComments(sheet, sheetPath); err != nil {
				return err
			}

			sheetData, err := utils.MarshalXMLWithHeader(sheet.worksheet)
			if err != nil {
				return err
			}
			if _, err := w.pkg.AddPart(sheetPath, packaging.ContentTypeWorksheet, sheetData); err != nil {
				return err
			}

			// Add relationship with the correct ID
			rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
			rels.AddWithID(sheet.relID, packaging.RelTypeWorksheet, "worksheets/sheet"+fmt.Sprintf("%d.xml", i+1), packaging.TargetModeInternal)
		}

	// Save shared strings
	if w.sharedStrings.Count() > 0 {
		ssData, err := w.sharedStrings.marshal()
		if err != nil {
			return err
		}
		if _, err := w.pkg.AddPart(packaging.ExcelSharedStringsPath, packaging.ContentTypeSharedStrings, ssData); err != nil {
			return err
		}
		rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
		relID := rels.NextID()
		if existing := rels.FirstByType(packaging.RelTypeSharedStrings); existing != nil {
			relID = existing.ID
		}
		rels.AddWithID(relID, packaging.RelTypeSharedStrings, "sharedStrings.xml", packaging.TargetModeInternal)
	}

	// Save styles
	if w.styles != nil && w.styles.stylesheet != nil {
		stylesData, err := utils.MarshalXMLWithHeader(w.styles.stylesheet)
		if err != nil {
			return err
		}
		if _, err := w.pkg.AddPart(packaging.ExcelStylesPath, packaging.ContentTypeExcelStyles, stylesData); err != nil {
			return err
		}
		rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
		relID := rels.NextID()
		if existing := rels.FirstByType(packaging.RelTypeStyles); existing != nil {
			relID = existing.ID
		}
		rels.AddWithID(relID, packaging.RelTypeStyles, "styles.xml", packaging.TargetModeInternal)
	}

	if err := w.writeAdvancedParts(); err != nil {
		return err
	}

	return nil
}

func (w *workbookImpl) writeAdvancedParts() error {
	if w.pkg == nil {
		return nil
	}
	for partPath, data := range w.themeParts {
		if _, err := w.pkg.AddPart(partPath, packaging.ContentTypeTheme, data); err != nil {
			return err
		}
		rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
		rel := rels.FirstByType(packaging.RelTypeTheme)
		relID := rels.NextID()
		if rel != nil {
			relID = rel.ID
		}
		rels.AddWithID(relID, packaging.RelTypeTheme, relativeTarget(packaging.ExcelWorkbookPath, partPath), packaging.TargetModeInternal)
	}
	for partPath, part := range w.extraParts {
		if part == nil {
			continue
		}
		content, err := part.Content()
		if err != nil {
			return err
		}
		if _, err := w.pkg.AddPart(partPath, part.ContentType(), content); err != nil {
			return err
		}
	}
	return nil
}

func (w *workbookImpl) updateComments(sheet *worksheetImpl, sheetPath string) error {
	if sheet.comments == nil {
		return nil
	}
	commentsPath := sheet.comments.path
	if commentsPath == "" {
		commentsPath = fmt.Sprintf("xl/comments/comment%d.xml", sheet.index+1)
		sheet.comments.path = commentsPath
	}
	data, err := utils.MarshalXMLWithHeader(sheet.comments.comments)
	if err != nil {
		return err
	}
	if _, err := w.pkg.AddPart(commentsPath, packaging.ContentTypeExcelComments, data); err != nil {
		return err
	}
	rels := w.pkg.GetRelationships(sheetPath)
	relID := sheet.comments.relID
	if relID == "" {
		relID = rels.NextID()
		sheet.comments.relID = relID
	}
	rels.AddWithID(relID, packaging.RelTypeComments, relativeTarget(sheetPath, commentsPath), packaging.TargetModeInternal)
	if err := w.updateCommentVML(sheet, sheetPath, rels); err != nil {
		return err
	}
	return nil
}

func (w *workbookImpl) updateCommentVML(sheet *worksheetImpl, sheetPath string, rels *packaging.Relationships) error {
	if sheet == nil || sheet.comments == nil {
		return nil
	}
	if sheet.comments.comments == nil || sheet.comments.comments.CommentList == nil {
		return nil
	}
	if len(sheet.comments.comments.CommentList.Comment) == 0 {
		return nil
	}
	if rels == nil {
		rels = w.pkg.GetRelationships(sheetPath)
	}
	vmlPath := sheet.comments.vmlPath
	if vmlPath == "" {
		vmlPath = fmt.Sprintf("xl/drawings/commentsDrawing%d.vml", sheet.index+1)
		sheet.comments.vmlPath = vmlPath
	}
	vmlData, err := buildCommentsVML(sheet.comments)
	if err != nil {
		return err
	}
	if _, err := w.pkg.AddPart(vmlPath, packaging.ContentTypeVML, vmlData); err != nil {
		return err
	}
	vmlRelID := sheet.comments.vmlRelID
	if vmlRelID == "" {
		vmlRelID = rels.NextID()
		sheet.comments.vmlRelID = vmlRelID
	}
	rels.AddWithID(vmlRelID, packaging.RelTypeVML, relativeTarget(sheetPath, vmlPath), packaging.TargetModeInternal)
	if sheet.worksheet.LegacyDrawing == nil {
		sheet.worksheet.LegacyDrawing = &sml.LegacyDrawing{}
	}
	if sheet.worksheet.LegacyDrawing.ID == "" {
		sheet.worksheet.LegacyDrawing.ID = vmlRelID
	}
	return nil
}

func relativeTarget(source, target string) string {
	if source == "" {
		return target
	}
	sourceDir := path.Dir(source)
	rel, err := filepath.Rel(sourceDir, target)
	if err != nil {
		return strings.TrimPrefix(packaging.ResolveRelationshipTarget(source, target), "xl/")
	}
	return filepath.ToSlash(rel)
}

func (w *workbookImpl) updateTables(sheet *worksheetImpl, sheetIndex int) error {
	if len(sheet.tables) == 0 {
		return nil
	}

	sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", sheetIndex)
	rels := w.pkg.GetRelationships(sheetPath)

	for i, table := range sheet.tables {
		tablePath := fmt.Sprintf("xl/tables/table%d.xml", i+1)
		table.path = tablePath
		tableData, err := utils.MarshalXMLWithHeader(table.table)
		if err != nil {
			return err
		}
		if _, err := w.pkg.AddPart(tablePath, packaging.ContentTypeTable, tableData); err != nil {
			return err
		}
		relID := table.relID
		if relID == "" {
			relID = rels.NextID()
			table.relID = relID
		}
		rels.AddWithID(relID, packaging.RelTypeTable, "../tables/table"+fmt.Sprintf("%d.xml", i+1), packaging.TargetModeInternal)
		if table.table.ID == 0 {
			table.table.ID = w.nextTableID
			w.nextTableID++
		}
	}
	return nil
}

// getSharedString returns the shared string at the given index.
func (w *workbookImpl) getSharedString(index int) string {
	return w.sharedStrings.Get(index)
}

// addSharedString adds a string to the shared strings table and returns its index.
func (w *workbookImpl) addSharedString(s string) int {
	return w.sharedStrings.Add(s)
}
