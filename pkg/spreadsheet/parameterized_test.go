package spreadsheet

import (
	"testing"
	"time"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// =============================================================================
// Test Fixtures
// =============================================================================

// WorkbookFixture represents a test fixture for workbooks.
type WorkbookFixture struct {
	Name        string
	Description string
	Setup       func(*Workbook)
}

// CommonFixtures provides standard workbook test fixtures.
var CommonFixtures = []WorkbookFixture{
	{
		Name:        "empty",
		Description: "Empty workbook with default sheet",
		Setup:       func(w *Workbook) {},
	},
	{
		Name:        "multiple_sheets",
		Description: "Workbook with three sheets",
		Setup: func(w *Workbook) {
			w.AddSheet("Sheet2")
			w.AddSheet("Sheet3")
		},
	},
	{
		Name:        "single_cell",
		Description: "Single cell with text",
		Setup: func(w *Workbook) {
			w.Sheets()[0].Cell("A1").SetValue("Hello")
		},
	},
	{
		Name:        "data_types",
		Description: "Various data types",
		Setup: func(w *Workbook) {
			sheet := w.Sheets()[0]
			sheet.Cell("A1").SetValue("Text")
			sheet.Cell("A2").SetValue(42)
			sheet.Cell("A3").SetValue(3.14159)
			sheet.Cell("A4").SetValue(true)
			sheet.Cell("A5").SetValue(time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC))
		},
	},
	{
		Name:        "formulas",
		Description: "Cells with formulas",
		Setup: func(w *Workbook) {
			sheet := w.Sheets()[0]
			sheet.Cell("A1").SetValue(10)
			sheet.Cell("A2").SetValue(20)
			sheet.Cell("A3").SetFormula("A1+A2")
			sheet.Cell("A4").SetFormula("SUM(A1:A2)")
		},
	},
	{
		Name:        "merged_cells",
		Description: "Sheet with merged cells",
		Setup: func(w *Workbook) {
			sheet := w.Sheets()[0]
			sheet.Cell("A1").SetValue("Merged Header")
			sheet.MergeCells("A1:C1")
		},
	},
	{
		Name:        "large_data",
		Description: "Sheet with 100 rows of data",
		Setup: func(w *Workbook) {
			sheet := w.Sheets()[0]
			for i := 1; i <= 100; i++ {
				sheet.CellByRC(i, 1).SetValue(i)
				sheet.CellByRC(i, 2).SetValue(i * 2)
				sheet.CellByRC(i, 3).SetValue(i * i)
			}
		},
	},
	{
		Name:        "hidden_sheet",
		Description: "Workbook with hidden sheet",
		Setup: func(w *Workbook) {
			w.AddSheet("Hidden")
			sheet, _ := w.Sheet("Hidden")
			sheet.SetHidden(true)
		},
	},
}

// =============================================================================
// Parameterized Fixture Tests
// =============================================================================

func TestFixtures_RoundTrip(t *testing.T) {
	for _, fixture := range CommonFixtures {
		t.Run(fixture.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, err := New()
			h.RequireNoError(err, "New()")
			
			fixture.Setup(w)
			
			sheetCount := w.SheetCount()
			path := h.TempFile(fixture.Name + ".xlsx")
			h.RequireNoError(w.SaveAs(path), "SaveAs()")
			w.Close()
			
			w2, err := Open(path)
			h.RequireNoError(err, "Open()")
			defer w2.Close()
			
			h.AssertEqual(w2.SheetCount(), sheetCount, "SheetCount after reload")
		})
	}
}

// =============================================================================
// Cell Value Type Parameterized Tests
// =============================================================================

type valueTypeTestCase struct {
	Name     string
	Value    interface{}
	WantType CellType
}

var valueTypeCases = []valueTypeTestCase{
	{"string", "Hello", CellTypeString},
	{"int", 42, CellTypeNumber},
	{"int64", int64(100), CellTypeNumber},
	{"float64", 3.14, CellTypeNumber},
	{"bool_true", true, CellTypeBoolean},
	{"bool_false", false, CellTypeBoolean},
	{"nil", nil, CellTypeEmpty},
}

func TestCellValueTypes_Parameterized(t *testing.T) {
	for _, tc := range valueTypeCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			cell := w.Sheets()[0].Cell("A1")
			cell.SetValue(tc.Value)
			
			h.AssertEqual(cell.Type(), tc.WantType, "Cell type")
		})
	}
}

// =============================================================================
// Cell Reference Parameterized Tests  
// =============================================================================

func TestCellReferences_Parameterized(t *testing.T) {
	for _, tc := range testutil.CommonCellRefCases {
		if tc.WantErr {
			continue // Skip error cases for this test
		}
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			cell := w.Sheets()[0].Cell(tc.Ref)
			if cell == nil {
				t.Fatalf("Cell(%s) returned nil", tc.Ref)
			}
			
			h.AssertEqual(cell.Row(), tc.Row, "Row")
			h.AssertEqual(cell.Column(), tc.Col, "Column")
		})
	}
}

// =============================================================================
// Range Parameterized Tests
// =============================================================================

func TestRanges_Parameterized(t *testing.T) {
	for _, tc := range testutil.CommonRangeCases {
		if tc.WantErr {
			continue
		}
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			rng := w.Sheets()[0].Range(tc.Ref)
			if rng == nil {
				t.Fatalf("Range(%s) returned nil", tc.Ref)
			}
			
			expectedRows := tc.EndRow - tc.StartRow + 1
			expectedCols := tc.EndCol - tc.StartCol + 1
			
			h.AssertEqual(rng.RowCount(), expectedRows, "RowCount")
			h.AssertEqual(rng.ColumnCount(), expectedCols, "ColumnCount")
		})
	}
}

// =============================================================================
// Numeric Value Parameterized Tests
// =============================================================================

func TestNumericValues_Parameterized(t *testing.T) {
	for _, tc := range testutil.CommonNumericCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			cell := w.Sheets()[0].Cell("A1")
			cell.SetValue(tc.Input)
			
			got, err := cell.Float64()
			h.RequireNoError(err, "Float64()")
			h.AssertEqual(got, tc.Want, "Float64 value")
		})
	}
}

// =============================================================================
// String Value Parameterized Tests
// =============================================================================

func TestStringValues_Parameterized(t *testing.T) {
	for _, tc := range testutil.CommonStringCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			cell := w.Sheets()[0].Cell("A1")
			cell.SetValue(tc.Input)
			
			got := cell.String()
			h.AssertEqual(got, tc.Want, "String value")
		})
	}
}

// =============================================================================
// String Values Round-Trip
// =============================================================================

func TestStringValues_RoundTrip(t *testing.T) {
	for _, tc := range testutil.CommonStringCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			w.Sheets()[0].Cell("A1").SetValue(tc.Input)
			
			path := h.TempFile(tc.Name + ".xlsx")
			h.RequireNoError(w.SaveAs(path), "SaveAs")
			w.Close()
			
			w2, _ := Open(path)
			defer w2.Close()
			
			got := w2.Sheets()[0].Cell("A1").String()
			h.AssertEqual(got, tc.Want, "String after round-trip")
		})
	}
}

// =============================================================================
// Sheet Operations Parameterized Tests
// =============================================================================

type sheetOpTestCase struct {
	Name         string
	InitialCount int // Not counting default Sheet1
	Operation    func(*Workbook)
	WantCount    int
}

var sheetOpCases = []sheetOpTestCase{
	{"add_one", 0, func(w *Workbook) { w.AddSheet("New") }, 2},
	{"add_multiple", 0, func(w *Workbook) { w.AddSheet("A"); w.AddSheet("B") }, 3},
	{"delete_by_name", 2, func(w *Workbook) { w.DeleteSheet("Sheet2") }, 2},
	{"delete_by_index", 2, func(w *Workbook) { w.DeleteSheet(1) }, 2},
}

func TestSheetOperations_Parameterized(t *testing.T) {
	for _, tc := range sheetOpCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			// Setup initial sheets
			for i := 0; i < tc.InitialCount; i++ {
				w.AddSheet("Sheet" + string(rune('2'+i)))
			}
			
			tc.Operation(w)
			
			h.AssertEqual(w.SheetCount(), tc.WantCount, "SheetCount")
		})
	}
}

// =============================================================================
// Formula Parameterized Tests
// =============================================================================

type formulaTestCase struct {
	Name    string
	Formula string
}

var formulaCases = []formulaTestCase{
	{"simple_add", "A1+B1"},
	{"simple_sub", "A1-B1"},
	{"simple_mul", "A1*B1"},
	{"simple_div", "A1/B1"},
	{"sum", "SUM(A1:A10)"},
	{"average", "AVERAGE(A1:A10)"},
	{"count", "COUNT(A1:A10)"},
	{"max", "MAX(A1:A10)"},
	{"min", "MIN(A1:A10)"},
	{"if", "IF(A1>0,\"Yes\",\"No\")"},
	{"nested", "SUM(A1:A10)/COUNT(A1:A10)"},
	{"cross_sheet", "Sheet2!A1+Sheet2!B1"},
}

func TestFormulas_Parameterized(t *testing.T) {
	for _, tc := range formulaCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			cell := w.Sheets()[0].Cell("C1")
			cell.SetFormula(tc.Formula)
			
			h.AssertEqual(cell.Formula(), tc.Formula, "Formula")
			h.AssertTrue(cell.HasFormula(), "HasFormula")
			h.AssertEqual(cell.Type(), CellTypeFormula, "Type")
		})
	}
}

// =============================================================================
// Row Height Parameterized Tests
// =============================================================================

type rowHeightTestCase struct {
	Name   string
	Height float64
}

var rowHeightCases = []rowHeightTestCase{
	{"small", 10},
	{"default", 15},
	{"medium", 20},
	{"large", 30},
	{"extra_large", 50},
}

func TestRowHeight_Parameterized(t *testing.T) {
	for _, tc := range rowHeightCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			row := w.Sheets()[0].Row(1)
			row.SetHeight(tc.Height)
			
			h.AssertEqual(row.Height(), tc.Height, "Row height")
		})
	}
}

// =============================================================================
// Merge Cell Parameterized Tests
// =============================================================================

type mergeCellTestCase struct {
	Name string
	Ref  string
}

var mergeCellCases = []mergeCellTestCase{
	{"row_merge", "A1:D1"},
	{"column_merge", "A1:A5"},
	{"block_merge", "A1:C3"},
	{"single_row_block", "B2:E2"},
	{"large_merge", "A1:Z10"},
}

func TestMergeCells_Parameterized(t *testing.T) {
	for _, tc := range mergeCellCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			sheet := w.Sheets()[0]
			sheet.MergeCells(tc.Ref)
			
			merged := sheet.MergedCells()
			h.AssertEqual(len(merged), 1, "Merged cell count")
			h.AssertEqual(merged[0], tc.Ref, "Merged cell reference")
		})
	}
}

// =============================================================================
// Date Value Parameterized Tests
// =============================================================================

type dateTestCase struct {
	Name string
	Date time.Time
}

var dateCases = []dateTestCase{
	{"epoch", time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)},
	{"y2k", time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC)},
	{"recent", time.Date(2024, 6, 15, 0, 0, 0, 0, time.UTC)},
	{"future", time.Date(2030, 12, 31, 0, 0, 0, 0, time.UTC)},
}

func TestDateValues_Parameterized(t *testing.T) {
	for _, tc := range dateCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			w, _ := New()
			defer w.Close()
			
			cell := w.Sheets()[0].Cell("A1")
			cell.SetValue(tc.Date)
			
			got, err := cell.Time()
			h.RequireNoError(err, "Time()")
			
			// Compare date parts (time may have precision issues)
			h.AssertEqual(got.Year(), tc.Date.Year(), "Year")
			h.AssertEqual(int(got.Month()), int(tc.Date.Month()), "Month")
			h.AssertEqual(got.Day(), tc.Date.Day(), "Day")
		})
	}
}
