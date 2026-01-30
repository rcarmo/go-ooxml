package spreadsheet

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
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
	sheet := w.Sheets()[0]
	if sheet.Name() != "Sheet1" {
		t.Errorf("First sheet name = %q, want %q", sheet.Name(), "Sheet1")
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
	if err := w.DeleteSheet(0); err == nil {
		t.Error("DeleteSheet() should error when deleting last sheet")
	}
}

func TestSheet(t *testing.T) {
	w := testutil.NewResource(t, New)

	w.AddSheet("MySheet")

	// Get by index
	sheet, err := w.Sheet(1)
	if err != nil {
		t.Fatalf("Sheet(1) error = %v", err)
	}
	if sheet.Name() != "MySheet" {
		t.Errorf("Name() = %q, want %q", sheet.Name(), "MySheet")
	}

	// Get by name
	sheet, err = w.Sheet("MySheet")
	if err != nil {
		t.Fatalf("Sheet(MySheet) error = %v", err)
	}
	if sheet.Index() != 1 {
		t.Errorf("Index() = %d, want 1", sheet.Index())
	}

	// Not found
	_, err = w.Sheet("NonExistent")
	if err != ErrSheetNotFound {
		t.Errorf("Sheet(NonExistent) error = %v, want ErrSheetNotFound", err)
	}
}

func TestSheetProperties(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.Sheets()[0]

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

	sheet := w.Sheets()[0]
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

	sheet := w.Sheets()[0]
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

	sheet := w.Sheets()[0]
	rng := sheet.Range("A1:B2")

	count := 0
	rng.ForEach(func(cell *Cell) error {
		count++
		return nil
	})

	if count != 4 {
		t.Errorf("ForEach visited %d cells, want 4", count)
	}
}

// =============================================================================
// Row Tests
// =============================================================================

func TestRow(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.Sheets()[0]
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

// =============================================================================
// Merged Cells Tests
// =============================================================================

func TestMergeCells(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.Sheets()[0]

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

	sheet := w.Sheets()[0]

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

	sheet := w.Sheets()[0]

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

	sheet := w.Sheets()[0]
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

	sheet1 := w.Sheets()[0]
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
	dataSheet, _ := w2.Sheet("Data")
	if dataSheet.Cell("A1").String() != "Name" {
		t.Errorf("A1 = %q, want %q", dataSheet.Cell("A1").String(), "Name")
	}

	// Verify formula
	summarySheet, _ := w2.Sheet("Summary")
	if summarySheet.Cell("A1").Formula() != "SUM(Data!B2)" {
		t.Errorf("Formula = %q, want %q", summarySheet.Cell("A1").Formula(), "SUM(Data!B2)")
	}
}

// =============================================================================
// Shared Strings Tests
// =============================================================================

func TestSharedStrings(t *testing.T) {
	w := testutil.NewResource(t, New)

	sheet := w.Sheets()[0]

	// Same string should share the same index
	sheet.Cell("A1").SetValue("Shared")
	sheet.Cell("A2").SetValue("Shared")
	sheet.Cell("A3").SetValue("Different")

	// Verify shared strings count (should be 2 unique strings)
	if w.sharedStrings.Count() != 2 {
		t.Errorf("SharedStrings count = %d, want 2", w.sharedStrings.Count())
	}
}
