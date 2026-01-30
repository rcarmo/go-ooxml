package spreadsheet

import (
	"os"
	"path/filepath"
	"testing"
	"time"
)

// =============================================================================
// Creation Tests
// =============================================================================

func TestNew(t *testing.T) {
	w, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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

func TestCellByReference(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]

	cell := sheet.Cell("A1")
	if cell == nil {
		t.Fatal("Cell(A1) returned nil")
	}

	if cell.Reference() != "A1" {
		t.Errorf("Reference() = %q, want %q", cell.Reference(), "A1")
	}

	if cell.Row() != 1 || cell.Column() != 1 {
		t.Errorf("Row/Col = (%d, %d), want (1, 1)", cell.Row(), cell.Column())
	}
}

func TestCellByRC(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]

	cell := sheet.CellByRC(5, 3) // E3
	if cell == nil {
		t.Fatal("CellByRC(5, 3) returned nil")
	}

	if cell.Reference() != "C5" {
		t.Errorf("Reference() = %q, want %q", cell.Reference(), "C5")
	}
}

func TestCellStringValue(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]
	cell := sheet.Cell("A1")

	cell.SetValue("Hello World")

	if cell.String() != "Hello World" {
		t.Errorf("String() = %q, want %q", cell.String(), "Hello World")
	}

	if cell.Type() != CellTypeString {
		t.Errorf("Type() = %d, want CellTypeString", cell.Type())
	}
}

func TestCellNumericValue(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]

	tests := []struct {
		name  string
		value interface{}
		want  float64
	}{
		{"int", 42, 42},
		{"int64", int64(100), 100},
		{"float64", 3.14159, 3.14159},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cell := sheet.Cell("B1")
			cell.SetValue(tc.value)

			got, err := cell.Float64()
			if err != nil {
				t.Fatalf("Float64() error = %v", err)
			}

			if got != tc.want {
				t.Errorf("Float64() = %f, want %f", got, tc.want)
			}

			if cell.Type() != CellTypeNumber {
				t.Errorf("Type() = %d, want CellTypeNumber", cell.Type())
			}
		})
	}
}

func TestCellBooleanValue(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]
	cell := sheet.Cell("A1")

	cell.SetValue(true)
	got, err := cell.Bool()
	if err != nil || !got {
		t.Errorf("Bool() = (%v, %v), want (true, nil)", got, err)
	}

	cell.SetValue(false)
	got, err = cell.Bool()
	if err != nil || got {
		t.Errorf("Bool() = (%v, %v), want (false, nil)", got, err)
	}

	if cell.Type() != CellTypeBoolean {
		t.Errorf("Type() = %d, want CellTypeBoolean", cell.Type())
	}
}

func TestCellDateValue(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]
	cell := sheet.Cell("A1")

	// Set a date
	testDate := time.Date(2024, 1, 15, 12, 30, 0, 0, time.UTC)
	cell.SetValue(testDate)

	got, err := cell.Time()
	if err != nil {
		t.Fatalf("Time() error = %v", err)
	}

	// Check date components (time may have slight differences due to float precision)
	if got.Year() != testDate.Year() || got.Month() != testDate.Month() || got.Day() != testDate.Day() {
		t.Errorf("Time() date = %v, want %v", got.Format("2006-01-02"), testDate.Format("2006-01-02"))
	}
}

func TestCellFormula(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]
	cell := sheet.Cell("C1")

	cell.SetFormula("A1+B1")

	if cell.Formula() != "A1+B1" {
		t.Errorf("Formula() = %q, want %q", cell.Formula(), "A1+B1")
	}

	if !cell.HasFormula() {
		t.Error("HasFormula() should be true")
	}

	if cell.Type() != CellTypeFormula {
		t.Errorf("Type() = %d, want CellTypeFormula", cell.Type())
	}
}

func TestCellClear(t *testing.T) {
	w, _ := New()
	defer w.Close()

	sheet := w.Sheets()[0]
	cell := sheet.Cell("A1")

	cell.SetValue("Test")
	cell.SetValue(nil)

	if cell.Type() != CellTypeEmpty {
		t.Errorf("Type() = %d, want CellTypeEmpty", cell.Type())
	}
}

// =============================================================================
// Range Tests
// =============================================================================

func TestRange(t *testing.T) {
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()
	defer w.Close()

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
	w, _ := New()

	sheet := w.Sheets()[0]
	sheet.Cell("A1").SetValue("Test")
	sheet.Cell("B1").SetValue(42)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "test.xlsx")

	if err := w.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	w.Close()

	// Verify file exists
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Error("Output file was not created")
	}
}

func TestRoundTrip(t *testing.T) {
	w, _ := New()

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
	w.Close()

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
	w, _ := New()
	defer w.Close()

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
