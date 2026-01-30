// Package spreadsheet provides a high-level API for working with Excel workbooks.
package spreadsheet

import (
	"fmt"
	"io"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Workbook represents an Excel workbook.
type Workbook struct {
	pkg           *packaging.Package
	workbook      *sml.Workbook
	sheets        []*Worksheet
	sharedStrings *SharedStrings
	path          string
	nextSheetID   int
}

// New creates a new empty workbook with one sheet.
func New() (*Workbook, error) {
	w := &Workbook{
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
		sheets:        make([]*Worksheet, 0),
		sharedStrings: newSharedStrings(),
		nextSheetID:   1,
	}

	if err := w.initPackage(); err != nil {
		return nil, err
	}

	// Add default sheet
	w.AddSheet("Sheet1")

	return w, nil
}

// Open opens an existing workbook from a file path.
func Open(path string) (*Workbook, error) {
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
func OpenReader(r io.ReaderAt, size int64) (*Workbook, error) {
	pkg, err := packaging.OpenReader(r, size)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

func openFromPackage(pkg *packaging.Package) (*Workbook, error) {
	w := &Workbook{
		pkg:           pkg,
		sheets:        make([]*Worksheet, 0),
		sharedStrings: newSharedStrings(),
		nextSheetID:   1,
	}

	// Parse workbook.xml
	if err := w.parseWorkbook(); err != nil {
		return nil, err
	}

	// Parse shared strings
	w.parseSharedStrings()

	// Parse worksheets
	if err := w.parseSheets(); err != nil {
		return nil, err
	}

	return w, nil
}

// Save saves the workbook to its original path.
func (w *Workbook) Save() error {
	if w.path == "" {
		return fmt.Errorf("no path set, use SaveAs")
	}
	return w.SaveAs(w.path)
}

// SaveAs saves the workbook to a new path.
func (w *Workbook) SaveAs(path string) error {
	if err := w.updatePackage(); err != nil {
		return err
	}
	return w.pkg.SaveAs(path)
}

// Close closes the workbook and releases resources.
func (w *Workbook) Close() error {
	return w.pkg.Close()
}

// =============================================================================
// Sheet access
// =============================================================================

// Sheets returns all worksheets in the workbook.
func (w *Workbook) Sheets() []*Worksheet {
	return w.sheets
}

// Sheet returns a worksheet by name or index.
func (w *Workbook) Sheet(nameOrIndex interface{}) (*Worksheet, error) {
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
func (w *Workbook) SheetCount() int {
	return len(w.sheets)
}

// AddSheet adds a new worksheet with the given name.
func (w *Workbook) AddSheet(name string) *Worksheet {
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

	sheet := &Worksheet{
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
func (w *Workbook) DeleteSheet(nameOrIndex interface{}) error {
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
		return fmt.Errorf("cannot delete the last sheet")
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

func (w *Workbook) initPackage() error {
	// Add main relationship
	w.pkg.AddRelationship("/", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
	return nil
}

func (w *Workbook) parseWorkbook() error {
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

func (w *Workbook) parseSharedStrings() {
	part, err := w.pkg.GetPart(packaging.ExcelSharedStringsPath)
	if err != nil {
		return // Shared strings are optional
	}
	data, err := part.Content()
	if err != nil {
		return
	}

	w.sharedStrings.parse(data)
}

func (w *Workbook) parseSheets() error {
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
				sheetPath = "xl/" + rel.Target
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

		sheet := &Worksheet{
			workbook:  w,
			worksheet: worksheet,
			name:      sheetRef.Name,
			sheetID:   sheetRef.SheetID,
			relID:     sheetRef.ID,
			index:     i,
			path:      sheetPath,
		}

		w.sheets = append(w.sheets, sheet)

		if sheetRef.SheetID >= w.nextSheetID {
			w.nextSheetID = sheetRef.SheetID + 1
		}
	}

	return nil
}

func (w *Workbook) updatePackage() error {
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
		w.pkg.AddRelationship(packaging.ExcelWorkbookPath, "sharedStrings.xml", packaging.RelTypeSharedStrings)
	}

	return nil
}

// getSharedString returns the shared string at the given index.
func (w *Workbook) getSharedString(index int) string {
	return w.sharedStrings.Get(index)
}

// addSharedString adds a string to the shared strings table and returns its index.
func (w *Workbook) addSharedString(s string) int {
	return w.sharedStrings.Add(s)
}
