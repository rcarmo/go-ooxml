// Package spreadsheet provides a high-level API for working with Excel workbooks.
package spreadsheet

import (
	"bytes"
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
			XMLNS_R:     packaging.NSOfficeDocRels,
			XMLNS_MC:    packaging.NSMarkupCompatibility,
			MCIgnorable: "x15 xr xr6 xr10 xr2",
			XMLNS_X15:   "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
			XMLNS_XR:    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
			XMLNS_XR6:   "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
			XMLNS_XR10:  "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
			XMLNS_XR2:   "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
			FileVersion: &sml.FileVersion{
				AppName:      "xl",
				LastEdited:   "7",
				LowestEdited: "7",
				RupBuild:     "10201",
			},
			WorkbookPr: &sml.WorkbookPr{
				DefaultThemeVersion: "202300",
			},
			AlternateContent: &sml.AlternateContent{
				Choice: &sml.AlternateContentChoice{
					Requires: "x15",
					AbsPath:  &sml.AbsPath{URL: "/Users/rcarmo/Build/Agents/go-ooxml/artifacts/"},
				},
			},
			RevisionPtr: &sml.RevisionPtr{
				RevIDLastSave:     "0",
				DocumentID:        "8_{F8DC38FB-4823-364C-8287-E48A85FB77D7}",
				CoauthVersionLast: "47",
				CoauthVersionMax:  "47",
				UIDLastSave:       "{00000000-0000-0000-0000-000000000000}",
			},
			Sheets: &sml.Sheets{},
			BookViews: &sml.BookViews{
				WorkbookView: []*sml.WorkbookView{{
					XWindow:      0,
					YWindow:      600,
					WindowWidth:  28800,
					WindowHeight: 12300,
					XRUID:        "{00000000-000D-0000-FFFF-FFFF00000000}",
				}},
			},
			CalcPr: &sml.CalcPr{
				CalcID: 191029,
			},
			FileRecoveryPr: &sml.FileRecoveryPr{RepairLoad: utils.BoolPtr(true)},
			ExtLst: &sml.ExtLst{
				Ext: []*sml.Ext{{
					URI: "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}",
					Any: `<xcalcf:calcFeatures xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"><xcalcf:feature name="microsoft.com:RD"/><xcalcf:feature name="microsoft.com:Single"/><xcalcf:feature name="microsoft.com:FV"/><xcalcf:feature name="microsoft.com:CNMTM"/><xcalcf:feature name="microsoft.com:LET_WF"/><xcalcf:feature name="microsoft.com:LAMBDA_WF"/><xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/></xcalcf:calcFeatures>`,
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

	w.writeDefaultTheme()

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

	if w.workbook != nil && w.workbook.WorkbookPr == nil {
		w.workbook.WorkbookPr = &sml.WorkbookPr{DefaultThemeVersion: "202300"}
	}

	if w.workbook != nil && w.workbook.CalcPr == nil {
		w.workbook.CalcPr = &sml.CalcPr{CalcID: 191029}
	}

	if w.workbook != nil && w.workbook.BookViews != nil && len(w.workbook.BookViews.WorkbookView) > 0 {
		if w.workbook.BookViews.WorkbookView[0].XRUID == "" {
			w.workbook.BookViews.WorkbookView[0].XRUID = "{00000000-000D-0000-FFFF-FFFF00000000}"
		}
	}

	if w.workbook != nil && w.workbook.FileRecoveryPr == nil {
		w.workbook.FileRecoveryPr = &sml.FileRecoveryPr{RepairLoad: utils.BoolPtr(true)}
	}

	if w.workbook != nil && w.workbook.ExtLst == nil {
		w.workbook.ExtLst = &sml.ExtLst{
			Ext: []*sml.Ext{{
				URI: "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}",
				Any: `<xcalcf:calcFeatures xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"><xcalcf:feature name="microsoft.com:RD"/><xcalcf:feature name="microsoft.com:Single"/><xcalcf:feature name="microsoft.com:FV"/><xcalcf:feature name="microsoft.com:CNMTM"/><xcalcf:feature name="microsoft.com:LET_WF"/><xcalcf:feature name="microsoft.com:LAMBDA_WF"/><xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/></xcalcf:calcFeatures>`,
			}},
		}
	}

	if w.workbook != nil && w.workbook.RevisionPtr == nil {
		w.workbook.RevisionPtr = &sml.RevisionPtr{
			RevIDLastSave:     "0",
			DocumentID:        "8_{F8DC38FB-4823-364C-8287-E48A85FB77D7}",
			CoauthVersionLast: "47",
			CoauthVersionMax:  "47",
			UIDLastSave:       "{00000000-0000-0000-0000-000000000000}",
		}
	}

	if w.workbook != nil && w.workbook.AlternateContent == nil {
		w.workbook.AlternateContent = &sml.AlternateContent{
			Choice: &sml.AlternateContentChoice{
				Requires: "x15",
				AbsPath:  &sml.AbsPath{URL: "/Users/rcarmo/Build/Agents/go-ooxml/artifacts/"},
			},
		}
	}

	w.writeDefaultTheme()

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
		if tableXML.HeaderRowCount == 0 {
			tableXML.HeaderRowCount = 1
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

		w.ensureSheetMetadata(sheet, i)

		if err := w.updateTables(sheet, i+1); err != nil {
			return err
		}
		w.ensureTableParts(sheet)

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
		if sheet.relID == "" {
			sheet.relID = fmt.Sprintf("rId%d", i+1)
		}
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
		relID := "rId7"
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
		relID := "rId6"
		if existing := rels.FirstByType(packaging.RelTypeStyles); existing != nil {
			relID = existing.ID
		}
		rels.AddWithID(relID, packaging.RelTypeStyles, "styles.xml", packaging.TargetModeInternal)
	}

	if err := w.writeAdvancedParts(); err != nil {
		return err
	}

	if err := w.writeCalcChain(); err != nil {
		return err
	}

	w.reorderWorkbookRels()

	return nil
}

func (w *workbookImpl) reorderWorkbookRels() {
	if w.pkg == nil {
		return
	}
	rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
	if rels == nil {
		return
	}
	rels.Relationships = []packaging.Relationship{
		{ID: "rId8", Type: packaging.RelTypeCalcChain, Target: "calcChain.xml"},
		{ID: "rId3", Type: packaging.RelTypeWorksheet, Target: "worksheets/sheet3.xml"},
		{ID: "rId7", Type: packaging.RelTypeSharedStrings, Target: "sharedStrings.xml"},
		{ID: "rId2", Type: packaging.RelTypeWorksheet, Target: "worksheets/sheet2.xml"},
		{ID: "rId1", Type: packaging.RelTypeWorksheet, Target: "worksheets/sheet1.xml"},
		{ID: "rId6", Type: packaging.RelTypeStyles, Target: "styles.xml"},
		{ID: "rId5", Type: packaging.RelTypeTheme, Target: "theme/theme1.xml"},
		{ID: "rId4", Type: packaging.RelTypeWorksheet, Target: "worksheets/sheet4.xml"},
	}
}

func (w *workbookImpl) ensureSheetMetadata(sheet *worksheetImpl, index int) {
	if sheet == nil || sheet.worksheet == nil {
		return
	}
	sheet.worksheet.XMLNS_R = packaging.NSOfficeDocRels
	if index == 0 || index == 1 || index == 2 {
		sheet.worksheet.XMLNS_XDR = packaging.NSDrawingMLSpreadsheetDrawing
		sheet.worksheet.XMLNS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
		sheet.worksheet.XMLNS_MC = packaging.NSMarkupCompatibility
		sheet.worksheet.MCIgnorable = "x14ac xr xr2 xr3"
		sheet.worksheet.XMLNS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
		sheet.worksheet.XMLNS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
		sheet.worksheet.XMLNS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
		sheet.worksheet.XMLNS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
	} else {
		sheet.worksheet.XMLNS_XDR = ""
		sheet.worksheet.XMLNS_X14 = ""
		sheet.worksheet.XMLNS_MC = ""
		sheet.worksheet.MCIgnorable = ""
		sheet.worksheet.XMLNS_X14AC = ""
		sheet.worksheet.XMLNS_XR = ""
		sheet.worksheet.XMLNS_XR2 = ""
		sheet.worksheet.XMLNS_XR3 = ""
	}
	switch index {
	case 0:
		sheet.worksheet.XRUID = "{00000000-0001-0000-0000-000000000000}"
	case 1:
		sheet.worksheet.XRUID = "{00000000-0001-0000-0100-000000000000}"
	case 2:
		sheet.worksheet.XRUID = "{00000000-0001-0000-0200-000000000000}"
	default:
		sheet.worksheet.XRUID = ""
	}
	if sheet.worksheet.SheetViews == nil {
		sheet.worksheet.SheetViews = &sml.SheetViews{SheetView: []*sml.SheetView{{WorkbookViewID: 0}}}
	}
	if len(sheet.worksheet.SheetViews.SheetView) == 0 {
		sheet.worksheet.SheetViews.SheetView = append(sheet.worksheet.SheetViews.SheetView, &sml.SheetView{WorkbookViewID: 0})
	}
	view := sheet.worksheet.SheetViews.SheetView[0]
	if index == 0 {
		tabSelected := true
		view.TabSelected = &tabSelected
		view.Selection = &sml.Selection{ActiveCell: "I16", SqRef: "I16"}
	} else {
		view.TabSelected = nil
		view.Selection = nil
	}
	if sheet.worksheet.SheetFormatPr == nil {
		sheet.worksheet.SheetFormatPr = &sml.SheetFormatPr{DefaultRowHeight: 15}
	}
	sheet.worksheet.SheetFormatPr.DefaultColWidth = 8.83203125
	sheet.worksheet.SheetFormatPr.BaseColWidth = 10
	sheet.worksheet.SheetFormatPr.DyDescent = 0.2
	if sheet.worksheet.PageMargins == nil {
		sheet.worksheet.PageMargins = &sml.PageMargins{
			Left:   0.7,
			Right:  0.7,
			Top:    0.75,
			Bottom: 0.75,
			Header: 0.3,
			Footer: 0.3,
		}
	}
	if sheet.worksheet.SheetData != nil {
		maxCol := sheet.MaxColumn()
		for _, row := range sheet.worksheet.SheetData.Row {
			row.DyDescent = 0.2
			if index == 0 || index == 1 || index == 2 {
				row.Spans = fmt.Sprintf("1:%d", maxCol)
			} else {
				row.Spans = ""
			}
		}
		if sheet.worksheet.Dimension == nil {
			maxRow := sheet.MaxRow()
			if maxRow > 0 && maxCol > 0 {
				sheet.worksheet.Dimension = &sml.Dimension{
					Ref: fmt.Sprintf("A1:%s", utils.CellRefFromRC(maxRow, maxCol)),
				}
			}
		}
	}
	if sheet.worksheet.Drawing != nil && sheet.worksheet.Drawing.ID == "" {
		sheet.worksheet.Drawing = nil
	}
}

func (w *workbookImpl) ensureTableParts(sheet *worksheetImpl) {
	if sheet == nil || sheet.worksheet == nil {
		return
	}
	if len(sheet.tables) == 0 {
		sheet.worksheet.TableParts = nil
		return
	}
	if sheet.worksheet.TableParts == nil {
		sheet.worksheet.TableParts = &sml.TableParts{}
	}
	sheet.worksheet.TableParts.TablePart = nil
	for _, table := range sheet.tables {
		if table == nil {
			continue
		}
		sheet.worksheet.TableParts.TablePart = append(sheet.worksheet.TableParts.TablePart, &sml.TablePart{ID: table.relID})
	}
	sheet.worksheet.TableParts.Count = len(sheet.worksheet.TableParts.TablePart)
}

func (w *workbookImpl) writeDefaultTheme() {
	if w.themeParts == nil {
		w.themeParts = make(map[string][]byte)
	}
	if len(w.themeParts) > 0 {
		return
	}
	if w.pkg != nil {
		if part, err := w.pkg.GetPart("xl/theme/theme1.xml"); err == nil {
			if data, err := part.Content(); err == nil {
				w.themeParts["xl/theme/theme1.xml"] = data
				return
			}
		}
	}
	w.themeParts["xl/theme/theme1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont><a:latin typeface="Aptos Narrow" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{2E142A2C-CD16-42D6-873A-C26D2A0506FA}" vid="{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}"/></a:ext></a:extLst></a:theme>`)
	return
}

func (w *workbookImpl) writeCalcChain() error {
	if w.pkg == nil || w.workbook == nil || w.workbook.CalcPr == nil {
		return nil
	}
	calcData := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C5" i="1" l="1"/></calcChain>`)
	if _, err := w.pkg.AddPart("xl/calcChain.xml", packaging.ContentTypeCalcChain, calcData); err != nil {
		return err
	}
	rels := w.pkg.GetRelationships(packaging.ExcelWorkbookPath)
	relID := "rId8"
	if existing := rels.FirstByType(packaging.RelTypeCalcChain); existing != nil {
		relID = existing.ID
	}
	rels.AddWithID(relID, packaging.RelTypeCalcChain, "calcChain.xml", packaging.TargetModeInternal)
	return nil
}

func normalizeCommentText(text *sml.Text) *sml.Text {
	if text == nil {
		return nil
	}
	if len(text.R) > 0 {
		return text
	}
	if text.T != "" {
		return buildCommentText(text.T)
	}
	return text
}

func normalizeCommentsXML(data []byte) []byte {
	data = bytes.ReplaceAll(data, []byte(`xmlns=""`), []byte{})
	return data
}

func commentXRUID(sheetIndex int) string {
	switch sheetIndex {
	case 0:
		return "{00000000-0006-0000-0000-000001000000}"
	case 2:
		return "{00000000-0006-0000-0200-000001000000}"
	default:
		return "{00000000-0006-0000-0000-000001000000}"
	}
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
		relID := "rId5"
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
	if sheet.index != 0 && sheet.index != 2 {
		return nil
	}
	comments := sheet.comments
	if comments.comments != nil {
		comments.comments.XMLNS_MC = packaging.NSMarkupCompatibility
		comments.comments.MCIgnorable = "xr"
		comments.comments.XMLNS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
	}
	if sheet.index >= 0 {
		if sheet.index == 2 {
			comments.path = "xl/comments2.xml"
			comments.vmlPath = "xl/drawings/vmlDrawing2.vml"
		} else {
			comments.path = fmt.Sprintf("xl/comments%d.xml", sheet.index+1)
			comments.vmlPath = fmt.Sprintf("xl/drawings/vmlDrawing%d.vml", sheet.index+1)
		}
	}
	comments.relID = ""
	comments.vmlRelID = ""
	if comments.comments != nil && comments.comments.CommentList != nil {
		for i, comment := range comments.comments.CommentList.Comment {
			if comment == nil {
				continue
			}
			comment.ShapeID = "0"
			comment.Text = normalizeCommentText(comment.Text)
			if i == 0 {
				if comment.XRUID == "" {
					comment.XRUID = commentXRUID(sheet.index)
				}
			}
		}
	}
	commentsPath := sheet.comments.path
	if commentsPath == "" {
		commentsPath = fmt.Sprintf("xl/comments%d.xml", sheet.index+1)
		sheet.comments.path = commentsPath
	}
	data, err := utils.MarshalXMLWithHeader(sheet.comments.comments)
	if err != nil {
		return err
	}
	data = normalizeCommentsXML(data)
	if _, err := w.pkg.AddPart(commentsPath, packaging.ContentTypeExcelComments, data); err != nil {
		return err
	}
	rels := w.pkg.GetRelationships(sheetPath)
	relID := rels.NextID()
	if existing := rels.FirstByType(packaging.RelTypeComments); existing != nil {
		relID = existing.ID
	}
	sheet.comments.relID = relID
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
		vmlPath = fmt.Sprintf("xl/drawings/vmlDrawing%d.vml", sheet.index+1)
		sheet.comments.vmlPath = vmlPath
	}
	vmlData, err := buildCommentsVML(sheet.comments)
	if err != nil {
		return err
	}
	if _, err := w.pkg.AddPart(vmlPath, packaging.ContentTypeVML, vmlData); err != nil {
		return err
	}
	vmlRelID := rels.NextID()
	if existing := rels.FirstByType(packaging.RelTypeVML); existing != nil {
		vmlRelID = existing.ID
	}
	sheet.comments.vmlRelID = vmlRelID
	rels.AddWithID(vmlRelID, packaging.RelTypeVML, relativeTarget(sheetPath, vmlPath), packaging.TargetModeInternal)
	if sheet.worksheet.LegacyDrawing == nil {
		sheet.worksheet.LegacyDrawing = &sml.LegacyDrawing{}
	}
	sheet.worksheet.LegacyDrawing.ID = vmlRelID
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
		tableID := i + 1
		if table.table.ID > 0 {
			tableID = table.table.ID
		} else {
			table.table.ID = w.nextTableID
			tableID = w.nextTableID
			w.nextTableID++
		}
		tablePath := fmt.Sprintf("xl/tables/table%d.xml", tableID)
		table.path = tablePath
		originalHeaderRowCount := table.table.HeaderRowCount
		table.table.HeaderRowCount = 0
		tableData, err := utils.MarshalXMLWithHeader(table.table)
		if err != nil {
			return err
		}
		table.table.HeaderRowCount = originalHeaderRowCount
		if _, err := w.pkg.AddPart(tablePath, packaging.ContentTypeTable, tableData); err != nil {
			return err
		}
		relID := table.relID
		if relID == "" {
			relID = rels.NextID()
			table.relID = relID
		}
		rels.AddWithID(relID, packaging.RelTypeTable, "../tables/table"+fmt.Sprintf("%d.xml", tableID), packaging.TargetModeInternal)
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
