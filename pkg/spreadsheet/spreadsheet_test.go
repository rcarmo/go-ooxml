package spreadsheet

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// =============================================================================
// Creation Tests
// =============================================================================

func TestNew(t *testing.T) {
	w := testutil.NewResource(t, New)

	// New workbook should have one sheet
	if got := w.SheetCount(); got != 1 {
		t.Errorf("SheetCount() = %d, want 1", got)
	}

	// First sheet should be named "Sheet1"
	sheet := w.SheetsRaw()[0]
	if sheet.Name() != "Sheet1" {
		t.Errorf("First sheet name = %q, want %q", sheet.Name(), "Sheet1")
	}
}

func TestWorkbook_CoreProperties(t *testing.T) {
	w := testutil.NewResource(t, New)

	props := &common.CoreProperties{
		Title:   "Workbook Title",
		Creator: "Workbook Author",
	}
	if err := w.SetCoreProperties(props); err != nil {
		t.Fatalf("SetCoreProperties() error = %v", err)
	}
	got, err := w.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties() error = %v", err)
	}
	if got.Title != props.Title {
		t.Errorf("Title = %q, want %q", got.Title, props.Title)
	}
	if got.Creator != props.Creator {
		t.Errorf("Creator = %q, want %q", got.Creator, props.Creator)
	}
}

func TestWorkbook_NamedRanges(t *testing.T) {
	w := testutil.NewResource(t, New)

	rng := w.AddNamedRange("SummaryTotal", "'Sheet1'!$B$2:$B$5")
	if rng.Name() != "SummaryTotal" {
		t.Errorf("Name() = %q, want %q", rng.Name(), "SummaryTotal")
	}
	if rng.RefersTo() != "'Sheet1'!$B$2:$B$5" {
		t.Errorf("RefersTo() = %q, want %q", rng.RefersTo(), "'Sheet1'!$B$2:$B$5")
	}
	if len(w.NamedRanges()) != 1 {
		t.Errorf("NamedRanges() count = %d, want 1", len(w.NamedRanges()))
	}

	rng.SetRefersTo("'Sheet1'!$B$2:$B$6")
	if rng.RefersTo() != "'Sheet1'!$B$2:$B$6" {
		t.Errorf("RefersTo() after update = %q, want %q", rng.RefersTo(), "'Sheet1'!$B$2:$B$6")
	}

	if rng.Hidden() {
		t.Error("Hidden() should default to false")
	}
	rng.SetHidden(true)
	if !rng.Hidden() {
		t.Error("Hidden() should be true after SetHidden(true)")
	}

	if _, ok := rng.SheetIndex(); ok {
		t.Error("SheetIndex() should be unset by default")
	}
	rng.SetSheetIndex(0)
	if index, ok := rng.SheetIndex(); !ok || index != 0 {
		t.Errorf("SheetIndex() = %d, %v; want 0, true", index, ok)
	}
	rng.ClearSheetIndex()
	if _, ok := rng.SheetIndex(); ok {
		t.Error("SheetIndex() should be cleared")
	}
}

func TestWorkbook_NamedRanges_RoundTrip(t *testing.T) {
	w := testutil.NewResource(t, New)
	w.AddNamedRange("Locations", "'Sheet1'!$A$2:$A$6")
	w.AddNamedRange("DaysSpent", "'Sheet1'!$B$2:$B$6")
	total := w.AddNamedRange("TotalDays", "'Sheet1'!$B$7")
	total.SetHidden(true)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "named-ranges.xlsx")
	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = w.Close()

	w2, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer w2.Close()

	got := w2.NamedRanges()
	if len(got) != 3 {
		t.Fatalf("NamedRanges() count = %d, want 3", len(got))
	}
	if got[0].Name() != "Locations" {
		t.Errorf("First name = %q, want %q", got[0].Name(), "Locations")
	}
	if got[1].RefersTo() != "'Sheet1'!$B$2:$B$6" {
		t.Errorf("DaysSpent RefersTo = %q, want %q", got[1].RefersTo(), "'Sheet1'!$B$2:$B$6")
	}
	if got[2].Hidden() != true {
		t.Error("TotalDays should be hidden after round-trip")
	}
}

func TestWorksheet_Comments(t *testing.T) {
	w := testutil.NewResource(t, New)
	sheet := w.SheetsRaw()[0]

	cell := sheet.Cell("A1")
	if cell == nil {
		t.Fatal("Cell(A1) returned nil")
	}

	if _, ok := cell.Comment(); ok {
		t.Error("Comment() should be empty for new cell")
	}

	if err := cell.SetComment("Test comment", "Test Author"); err != nil {
		t.Fatalf("SetComment() error = %v", err)
	}

	comment, ok := cell.Comment()
	if !ok {
		t.Fatal("Comment() should return comment after SetComment")
	}
	if comment.Text() != "Test comment" {
		t.Errorf("Text() = %q, want %q", comment.Text(), "Test comment")
	}
	if comment.Author() != "Test Author" {
		t.Errorf("Author() = %q, want %q", comment.Author(), "Test Author")
	}

	comments := sheet.Comments()
	if len(comments) != 1 {
		t.Fatalf("Comments() count = %d, want 1", len(comments))
	}

	comment.SetText("Updated")
	if comment.Text() != "Updated" {
		t.Errorf("Updated Text() = %q, want %q", comment.Text(), "Updated")
	}
}

func TestWorksheet_Comments_RoundTrip(t *testing.T) {
	w := testutil.NewResource(t, New)
	sheet := w.SheetsRaw()[0]

	_ = sheet.Cell("A2").SetComment("First comment", "Test Author")
	_ = sheet.Cell("B3").SetComment("Second comment", "Test Author")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "comments.xlsx")
	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = w.Close()

	w2, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer w2.Close()

	sheet2 := w2.SheetsRaw()[0]
	comments := sheet2.Comments()
	if len(comments) != 2 {
		t.Fatalf("Comments() count = %d, want 2", len(comments))
	}
	if comments[0].Reference() != "A2" {
		t.Errorf("First comment reference = %q, want %q", comments[0].Reference(), "A2")
	}
	if comments[1].Text() != "Second comment" {
		t.Errorf("Second comment text = %q, want %q", comments[1].Text(), "Second comment")
	}
}

func TestWorksheet_Relationships_WithTablesDrawingsAndComments(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet1 := w.SheetsRaw()[0]
	sheet1.Cell("A1").SetValue("H1")
	sheet1.Cell("B1").SetValue("H2")
	sheet1.Cell("C1").SetValue("H3")
	sheet1.Cell("A2").SetValue("R1")
	sheet1.Cell("B2").SetValue(1)
	sheet1.Cell("C2").SetValue(2)
	table := sheet1.AddTable("A1:C2", "OverviewTable")
	_ = table.UpdateRow(1, map[string]interface{}{
		"Column1": "R1",
		"Column2": 1,
		"Column3": 2,
	})
	if err := sheet1.AddChart("A4", "E12", "Chart"); err != nil {
		t.Fatalf("AddChart() error = %v", err)
	}
	if err := sheet1.AddDiagram("A14", "C20", "Diagram"); err != nil {
		t.Fatalf("AddDiagram() error = %v", err)
	}
	if err := sheet1.Cell("A2").SetComment("note", "author"); err != nil {
		t.Fatalf("SetComment() error = %v", err)
	}

	_ = w.AddSheet("Budget")
	risk := w.AddSheet("Risk")
	risk.Cell("A1").SetValue("Risk")
	risk.Cell("B1").SetValue("Level")
	risk.Cell("A2").SetValue("Exposure")
	risk.Cell("B2").SetValue("High")
	riskTable := risk.AddTable("A1:B2", "RiskTable")
	_ = riskTable.UpdateRow(1, map[string]interface{}{
		"Column1": "Exposure",
		"Column2": "High",
	})
	if err := risk.Cell("A2").SetComment("watch", "author"); err != nil {
		t.Fatalf("SetComment() error = %v", err)
	}

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "rels.xlsx")
	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = w.Close()

	round, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer round.Close()

	impl := round.(*workbookImpl)
	s1 := impl.sheets[0]
	rels1 := impl.pkg.GetRelationships("xl/worksheets/sheet1.xml")
	if s1.worksheet.Drawing == nil || s1.worksheet.Drawing.ID == "" {
		t.Fatal("sheet1 drawing relationship ID missing")
	}
	if s1.worksheet.LegacyDrawing == nil || s1.worksheet.LegacyDrawing.ID == "" {
		t.Fatal("sheet1 legacy drawing relationship ID missing")
	}
	if s1.worksheet.Drawing.ID == s1.worksheet.LegacyDrawing.ID {
		t.Fatalf("sheet1 drawing and legacy drawing share ID %q", s1.worksheet.Drawing.ID)
	}
	if rel := rels1.ByID(s1.worksheet.Drawing.ID); rel == nil || rel.Type != packaging.RelTypeDrawing {
		t.Fatalf("sheet1 drawing relationship invalid for %q", s1.worksheet.Drawing.ID)
	}
	if rel := rels1.ByID(s1.worksheet.LegacyDrawing.ID); rel == nil || rel.Type != packaging.RelTypeVML {
		t.Fatalf("sheet1 legacy drawing relationship invalid for %q", s1.worksheet.LegacyDrawing.ID)
	}
	for _, tp := range s1.worksheet.TableParts.TablePart {
		if rel := rels1.ByID(tp.ID); rel == nil || rel.Type != packaging.RelTypeTable {
			t.Fatalf("sheet1 table part relationship invalid for %q", tp.ID)
		}
	}

	s3 := impl.sheets[2]
	rels3 := impl.pkg.GetRelationships("xl/worksheets/sheet3.xml")
	for _, tp := range s3.worksheet.TableParts.TablePart {
		if rel := rels3.ByID(tp.ID); rel == nil || rel.Type != packaging.RelTypeTable {
			t.Fatalf("sheet3 table part relationship invalid for %q", tp.ID)
		}
	}
}

// =============================================================================
// Sheet Management Tests
// =============================================================================

func TestAddSheet(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet2 := w.AddSheet("Sheet2")
	if sheet2 == nil {
		t.Fatal("AddSheet() returned nil")
	}

	if w.SheetCount() != 2 {
		t.Errorf("SheetCount() = %d, want 2", w.SheetCount())
	}

	if sheet2.Name() != "Sheet2" {
		t.Errorf("Name() = %q, want %q", sheet2.Name(), "Sheet2")
	}

	if sheet2.Index() != 1 {
		t.Errorf("Index() = %d, want 1", sheet2.Index())
	}
}

func TestDeleteSheet(t *testing.T) {
	w := testutil.NewResource(t, New)

	w.AddSheet("Sheet2")
	w.AddSheet("Sheet3")

	// Delete by name
	if err := w.DeleteSheet("Sheet2"); err != nil {
		t.Fatalf("DeleteSheet(Sheet2) error = %v", err)
	}

	if w.SheetCount() != 2 {
		t.Errorf("SheetCount() = %d, want 2", w.SheetCount())
	}

	// Delete by index
	if err := w.DeleteSheet(0); err != nil {
		t.Fatalf("DeleteSheet(0) error = %v", err)
	}

	if w.SheetCount() != 1 {
		t.Errorf("SheetCount() = %d, want 1", w.SheetCount())
	}

	// Cannot delete last sheet
	if err := w.DeleteSheet(0); err != ErrCannotDeleteLastSheet {
		t.Errorf("DeleteSheet(last) error = %v, want ErrCannotDeleteLastSheet", err)
	}
}

func TestSheet(t *testing.T) {
	w := testutil.NewResource(t, New)

	w.AddSheet("MySheet")

	// Get by index
	sheet, err := w.SheetRaw(1)
	if err != nil {
		t.Fatalf("Sheet(1) error = %v", err)
	}
	if sheet.Name() != "MySheet" {
		t.Errorf("Name() = %q, want %q", sheet.Name(), "MySheet")
	}

	// Get by name
	sheet, err = w.SheetRaw("MySheet")
	if err != nil {
		t.Fatalf("Sheet(MySheet) error = %v", err)
	}
	if sheet.Index() != 1 {
		t.Errorf("Index() = %d, want 1", sheet.Index())
	}

	// Not found
	_, err = w.SheetRaw("NonExistent")
	if err != ErrSheetNotFound {
		t.Errorf("Sheet(NonExistent) error = %v, want ErrSheetNotFound", err)
	}
}

func TestSheetProperties(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]

	// Set name
	sheet.SetName("Renamed")
	if sheet.Name() != "Renamed" {
		t.Errorf("Name() = %q, want %q", sheet.Name(), "Renamed")
	}

	// Hidden
	if sheet.Hidden() {
		t.Error("New sheet should not be hidden")
	}

	sheet.SetHidden(true)
	if !sheet.Hidden() {
		t.Error("Sheet should be hidden after SetHidden(true)")
	}

	sheet.SetVisible(true)
	if !sheet.Visible() {
		t.Error("Sheet should be visible after SetVisible(true)")
	}
}

// =============================================================================
// Cell Tests
// =============================================================================

// Cell access/value tests live in parameterized_test.go

// =============================================================================
// Range Tests
// =============================================================================

func TestRange(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]
	rng := sheet.Range("A1:C3")

	if rng == nil {
		t.Fatal("Range(A1:C3) returned nil")
	}

	if rng.Reference() != "A1:C3" {
		t.Errorf("Reference() = %q, want %q", rng.Reference(), "A1:C3")
	}

	if rng.RowCount() != 3 {
		t.Errorf("RowCount() = %d, want 3", rng.RowCount())
	}

	if rng.ColumnCount() != 3 {
		t.Errorf("ColumnCount() = %d, want 3", rng.ColumnCount())
	}
}

func TestRangeSetValue(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]
	rng := sheet.Range("A1:B2")

	rng.SetValue("Test")

	// All cells should have "Test"
	for _, row := range rng.Cells() {
		for _, cell := range row {
			if cell.String() != "Test" {
				t.Errorf("Cell %s = %q, want %q", cell.Reference(), cell.String(), "Test")
			}
		}
	}
}

func TestRangeForEach(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]
	rng := sheet.Range("A1:B2")

	count := 0
	rng.ForEach(func(cell Cell) error {
		count++
		return nil
	})

	if count != 4 {
		t.Errorf("ForEach visited %d cells, want 4", count)
	}
}

func TestCellStyleFormatting(t *testing.T) {
	w := testutil.NewResource(t, New)
	sheet := w.SheetsRaw()[0]

	style := w.Styles().Style().
		SetFontName("Arial").
		SetFontSize(12).
		SetBold(true).
		SetItalic(true).
		SetFillColor("00FF0000").
		SetBorder(Border{Style: "thin"}).
		SetHorizontalAlignment(Alignment("center")).
		SetVerticalAlignment(Alignment("center")).
		SetNumberFormat("0.00")

	cell := sheet.Cell("A1")
	if err := cell.SetStyle(style); err != nil {
		t.Fatalf("SetStyle() error = %v", err)
	}

	if cell.NumberFormat() != "0.00" {
		t.Errorf("NumberFormat() = %q, want %q", cell.NumberFormat(), "0.00")
	}

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "styles.xlsx")
	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = w.Close()

	w2, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer w2.Close()

	cell2 := w2.SheetsRaw()[0].Cell("A1")
	if cell2.NumberFormat() != "0.00" {
		t.Errorf("NumberFormat() after reload = %q, want %q", cell2.NumberFormat(), "0.00")
	}
}

// =============================================================================
// Row Tests
// =============================================================================

func TestRow(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]
	row := sheet.Row(5)

	if row.Index() != 5 {
		t.Errorf("Index() = %d, want 5", row.Index())
	}

	row.SetHeight(30)
	if row.Height() != 30 {
		t.Errorf("Height() = %f, want 30", row.Height())
	}

	row.SetHidden(true)
	if !row.Hidden() {
		t.Error("Hidden() should be true")
	}
}

func TestRowIterator(t *testing.T) {
	w := testutil.NewResource(t, New)
	sheet := w.SheetsRaw()[0]
	sheet.Cell("A1").SetValue("first")
	sheet.Cell("A3").SetValue("third")

	iter := sheet.Rows()
	row, ok := iter.Next()
	if !ok || row.Index() != 1 {
		t.Fatalf("First row = %v, %v; want index 1", row, ok)
	}
	row, ok = iter.Next()
	if !ok || row.Index() != 3 {
		t.Fatalf("Second row = %v, %v; want index 3", row, ok)
	}
	if _, ok := iter.Next(); ok {
		t.Fatal("Expected iterator to finish")
	}
}

// =============================================================================
// Merged Cells Tests
// =============================================================================

func TestMergeCells(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]

	sheet.MergeCells("A1:C1")
	sheet.MergeCells("A2:A4")

	merged := sheet.MergedCells()
	if len(merged) != 2 {
		t.Errorf("MergedCells() count = %d, want 2", len(merged))
	}

	sheet.UnmergeCells("A1:C1")
	merged = sheet.MergedCells()
	if len(merged) != 1 {
		t.Errorf("After unmerge, count = %d, want 1", len(merged))
	}
}

// =============================================================================
// Dimensions Tests
// =============================================================================

func TestDimensions(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]

	// Empty sheet
	if sheet.MaxRow() != 0 || sheet.MaxColumn() != 0 {
		t.Errorf("Empty sheet dimensions = (%d, %d), want (0, 0)", sheet.MaxRow(), sheet.MaxColumn())
	}

	// Add some data
	sheet.Cell("C5").SetValue("test")
	sheet.Cell("E10").SetValue("test")

	if sheet.MaxRow() != 10 {
		t.Errorf("MaxRow() = %d, want 10", sheet.MaxRow())
	}

	if sheet.MaxColumn() != 5 {
		t.Errorf("MaxColumn() = %d, want 5", sheet.MaxColumn())
	}
}

func TestUsedRange(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]

	// Empty sheet
	if sheet.UsedRange() != nil {
		t.Error("Empty sheet should have nil UsedRange")
	}

	// Add data
	sheet.Cell("B2").SetValue("start")
	sheet.Cell("D4").SetValue("end")

	used := sheet.UsedRange()
	if used == nil {
		t.Fatal("UsedRange() returned nil")
	}
}

// =============================================================================
// Save/Load Tests
// =============================================================================

func TestSaveAs(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]
	sheet.Cell("A1").SetValue("Test")
	sheet.Cell("B1").SetValue(42)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "test.xlsx")

	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = w.Close()

	// Verify file exists
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Error("Output file was not created")
	}
}

func TestRoundTrip(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet1 := w.SheetsRaw()[0]
	sheet1.SetName("Data")
	sheet1.Cell("A1").SetValue("Name")
	sheet1.Cell("B1").SetValue("Age")
	sheet1.Cell("A2").SetValue("Alice")
	sheet1.Cell("B2").SetValue(30)

	sheet2 := w.AddSheet("Summary")
	sheet2.Cell("A1").SetFormula("SUM(Data!B2)")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "roundtrip.xlsx")

	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = w.Close()

	// Reopen
	w2, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer w2.Close()

	// Verify structure
	if w2.SheetCount() != 2 {
		t.Errorf("SheetCount() = %d, want 2", w2.SheetCount())
	}

	// Verify data
	dataSheet, _ := w2.SheetRaw("Data")
	if dataSheet.Cell("A1").String() != "Name" {
		t.Errorf("A1 = %q, want %q", dataSheet.Cell("A1").String(), "Name")
	}

	// Verify formula
	summarySheet, _ := w2.SheetRaw("Summary")
	if summarySheet.Cell("A1").Formula() != "SUM(Data!B2)" {
		t.Errorf("Formula = %q, want %q", summarySheet.Cell("A1").Formula(), "SUM(Data!B2)")
	}
}

func TestAdvancedPartsRoundTrip(t *testing.T) {
	orig := testutil.OpenResource(t, Open, filepath.Join("..", "..", "testdata", "excel", "formatting.xlsx"))

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "advanced-roundtrip.xlsx")
	if err := orig.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}

	round := testutil.OpenResource(t, Open, path)
	rels := round.(*workbookImpl).pkg.GetRelationships(packaging.ExcelWorkbookPath)
	if rels.FirstByType(packaging.RelTypeTheme) == nil {
		t.Error("Expected theme relationship after round-trip")
	}
}

func TestFixtureRoundTrip_Formatting(t *testing.T) {
	orig, err := Open("/workspace/testdata/excel/formatting.xlsx")
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer orig.Close()

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "formatting-roundtrip.xlsx")
	if err := orig.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}

	round, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer round.Close()

	cell := round.SheetsRaw()[0].Cell("A1")
	if cell.Style() == nil {
		t.Error("Style() should return a style for formatted cell")
	}
}

func TestFixtureRoundTrip_Formulas(t *testing.T) {
	orig, err := Open("/workspace/testdata/excel/formulas.xlsx")
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer orig.Close()

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "formulas-roundtrip.xlsx")
	if err := orig.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}

	round, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer round.Close()

	cell := round.SheetsRaw()[0].Cell("D2")
	if cell.Formula() != "B2*C2" {
		t.Errorf("Formula() = %q, want %q", cell.Formula(), "B2*C2")
	}
}

func TestFixtureRoundTrip_ConditionalFormatting(t *testing.T) {
	orig, err := Open("/workspace/testdata/excel/conditional_format.xlsx")
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer orig.Close()

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "conditional-format-roundtrip.xlsx")
	if err := orig.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}

	round, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer round.Close()

	if round.SheetsRaw()[0].worksheet == nil || round.SheetsRaw()[0].worksheet.SheetData == nil {
		t.Fatal("Sheet data missing after round-trip")
	}
	if round.SheetsRaw()[0].worksheet.ConditionalFormatting == nil {
		t.Fatal("ConditionalFormatting missing after round-trip")
	}
}

// =============================================================================
// Shared Strings Tests
// =============================================================================

func TestSharedStrings(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.SheetsRaw()[0]

	// Same string should share the same index
	sheet.Cell("A1").SetValue("Shared")
	sheet.Cell("A2").SetValue("Shared")
	sheet.Cell("A3").SetValue("Different")

	// Verify shared strings count (should be 2 unique strings)
	if w.SharedStrings().Count() != 2 {
		t.Errorf("SharedStrings count = %d, want 2", w.SharedStrings().Count())
	}
}
