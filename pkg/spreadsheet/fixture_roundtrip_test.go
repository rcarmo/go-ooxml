package spreadsheet

import (
	"path/filepath"
	"testing"
	"time"

)

type spreadsheetFixtureCase struct {
	name   string
	mutate func(t *testing.T, wb Workbook)
	verify func(t *testing.T, wb Workbook)
}

func TestFixtureRoundTrip_Complex(t *testing.T) {
	cases := []spreadsheetFixtureCase{
		{
			name: "minimal.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.Cell("A1").SetValue("Fixture")
				_ = sheet.Cell("B1").SetValue(42)
				_ = sheet.Cell("C1").SetFormula("B1*2")
				_ = sheet.MergeCells("A1:B1")
				_ = sheet.Cell("A1").SetComment("Fixture comment", "Tester")
				style := wb.Styles().Style().
					SetBold(true).
					SetFillColor("FFDDDD").
					SetHorizontalAlignment(Alignment("center")).
					SetVerticalAlignment(Alignment("center"))
				_ = sheet.Cell("A1").SetStyle(style)
				wb.AddNamedRange("FixtureRange", "'Sheet1'!$A$1:$B$1")
				sheet.Row(1).SetHeight(24)
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if sheet.Cell("A1").String() != "Fixture" {
					t.Errorf("A1 = %q, want Fixture", sheet.Cell("A1").String())
				}
				if sheet.Cell("C1").Formula() != "B1*2" {
					t.Errorf("C1 formula = %q, want B1*2", sheet.Cell("C1").Formula())
				}
				if !hasMergedRange(sheet, "A1:B1") {
					t.Error("Expected merged range A1:B1 after round-trip")
				}
				if comment, ok := sheet.Cell("A1").Comment(); !ok || comment.Text() != "Fixture comment" {
					t.Error("Expected comment after round-trip")
				}
				if !hasNamedRange(wb, "FixtureRange") {
					t.Error("Expected named range after round-trip")
				}
				if style := sheet.Cell("A1").Style(); style == nil {
					t.Error("Expected style after round-trip")
				}
				if hAlign, vAlign := cellAlignment(sheet, "A1"); hAlign != Alignment("center") || vAlign != Alignment("center") {
					t.Errorf("Alignment = (%q, %q), want (center, center)", hAlign, vAlign)
				}
				if sheet.Row(1).Height() != 24 {
					t.Errorf("Row 1 height = %v, want 24", sheet.Row(1).Height())
				}
			},
		},
		{
			name: "single_cell.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.Cell("A1").SetValue("Updated")
				_ = sheet.Cell("B1").SetValue(10.5)
				_ = sheet.Cell("C1").SetFormula("B1*3")
				_ = sheet.Cell("B1").SetNumberFormat("0.00")
				_ = sheet.Cell("D1").SetValue(1234.56)
				_ = sheet.Cell("D1").SetNumberFormat("$#,##0.00")
				_ = sheet.Cell("E1").SetValue(-987.65)
				_ = sheet.Cell("E1").SetNumberFormat("$#,##0.00_);[Red]($#,##0.00)")
				_ = sheet.Cell("F1").SetValue(0.25)
				_ = sheet.Cell("F1").SetNumberFormat("0.00%")
				_ = sheet.Cell("G1").SetValue(1000000)
				_ = sheet.Cell("G1").SetNumberFormat("#,##0")
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if sheet.Cell("A1").String() != "Updated" {
					t.Errorf("A1 = %q, want Updated", sheet.Cell("A1").String())
				}
				if sheet.Cell("C1").Formula() != "B1*3" {
					t.Errorf("C1 formula = %q, want B1*3", sheet.Cell("C1").Formula())
				}
				if sheet.Cell("B1").NumberFormat() != "0.00" {
					t.Errorf("B1 format = %q, want 0.00", sheet.Cell("B1").NumberFormat())
				}
				if sheet.Cell("D1").NumberFormat() != "$#,##0.00" {
					t.Errorf("D1 format = %q, want $#,##0.00", sheet.Cell("D1").NumberFormat())
				}
				if sheet.Cell("E1").NumberFormat() != "$#,##0.00_);[Red]($#,##0.00)" {
					t.Errorf("E1 format = %q, want $#,##0.00_);[Red]($#,##0.00)", sheet.Cell("E1").NumberFormat())
				}
				if sheet.Cell("F1").NumberFormat() != "0.00%" {
					t.Errorf("F1 format = %q, want 0.00%%", sheet.Cell("F1").NumberFormat())
				}
				if sheet.Cell("G1").NumberFormat() != "#,##0" {
					t.Errorf("G1 format = %q, want #,##0", sheet.Cell("G1").NumberFormat())
				}
			},
		},
		{
			name: "data_types.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.Cell("A1").SetValue("Fixture Text")
				_ = sheet.Cell("A2").SetValue(99)
				_ = sheet.Cell("A3").SetValue(true)
				_ = sheet.Cell("A4").SetValue(time.Date(2024, 2, 2, 0, 0, 0, 0, time.UTC))
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if sheet.Cell("A1").String() != "Fixture Text" {
					t.Errorf("A1 = %q, want Fixture Text", sheet.Cell("A1").String())
				}
				if got, err := sheet.Cell("A2").Int(); err != nil || got != 99 {
					t.Errorf("A2 = %v (err %v), want 99", got, err)
				}
				if v, _ := sheet.Cell("A3").Bool(); !v {
					t.Error("A3 expected true")
				}
				if _, err := sheet.Cell("A4").Time(); err != nil {
					t.Errorf("A4 time error = %v", err)
				}
			},
		},
		{
			name: "formatting.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				style := wb.Styles().Style().
					SetBold(true).
					SetFillColor("FF0000").
					SetBorder(Border{Style: "thin"}).
					SetHorizontalAlignment(Alignment("center")).
					SetVerticalAlignment(Alignment("center"))
				_ = sheet.Cell("A1").SetStyle(style)
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if style := sheet.Cell("A1").Style(); style == nil {
					t.Error("Expected formatted cell style after round-trip")
				}
				ws := sheet
				if ws == nil || ws.worksheet == nil || ws.worksheet.PageMargins == nil {
					t.Fatal("Expected page margins on formatting fixture")
				}
				margins := ws.worksheet.PageMargins
				if margins.Left != 0.75 || margins.Right != 0.75 || margins.Top != 1 || margins.Bottom != 1 || margins.Header != 0.5 || margins.Footer != 0.5 {
					t.Errorf("PageMargins = %#v, want left/right 0.75, top/bottom 1, header/footer 0.5", margins)
				}
				if hAlign, vAlign := cellAlignment(sheet, "A1"); hAlign != Alignment("center") || vAlign != Alignment("center") {
					t.Errorf("Alignment = (%q, %q), want (center, center)", hAlign, vAlign)
				}
			},
		},
		{
			name: "multiple_sheets.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.AddSheet("Fixture")
				sheet.SetHidden(true)
				_ = sheet.Cell("A1").SetValue("Hidden")
				wb.AddNamedRange("FixtureHidden", "'Fixture'!$A$1")
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet, err := wb.Sheet("Fixture")
				if err != nil {
					t.Fatalf("Sheet(Fixture) error = %v", err)
				}
				if !sheet.Hidden() {
					t.Error("Expected Fixture sheet hidden after round-trip")
				}
				if !hasNamedRange(wb, "FixtureHidden") {
					t.Error("Expected named range after round-trip")
				}
			},
		},
		{
			name: "tables.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				table := sheet.AddTable("A1:C3", "FixtureTable")
				_ = table.UpdateRow(1, map[string]interface{}{
					"Column1": "Item",
					"Column2": "Qty",
					"Column3": "Price",
				})
				_ = table.UpdateRow(2, map[string]interface{}{
					"Column1": "Widget",
					"Column2": 2,
					"Column3": 19.99,
				})
				_ = table.AddRow(map[string]interface{}{
					"Column1": "Gadget",
					"Column2": 5,
					"Column3": 3.5,
				})
				_ = table.DeleteRow(1)
				table, _ = wb.Table("FixtureTable")
				_ = table.UpdateRow(1, map[string]interface{}{
					"Column2": 7,
				})
				if cell := sheet.Cell("C2"); cell != nil {
					_ = cell.SetNumberFormat("$#,##0.00")
				}
			},
			verify: func(t *testing.T, wb Workbook) {
				table, err := wb.Table("FixtureTable")
				if err != nil {
					t.Fatalf("Table(FixtureTable) error = %v", err)
				}
				if len(table.Rows()) == 0 {
					t.Error("Expected table rows after round-trip")
				}
				if got := table.Headers(); len(got) != 3 || got[0] != "Column1" || got[1] != "Column2" || got[2] != "Column3" {
					t.Errorf("Table headers = %#v, want Column1/Column2/Column3", got)
				}
				if got := table.Reference(); got == "" {
					t.Error("Expected table reference after round-trip")
				}
				rows := table.Rows()
				if len(rows) < 2 {
					t.Fatalf("Expected at least 2 rows after round-trip, got %d", len(rows))
				}
				values := rows[0].Values()
				if values["Column2"] != 7.0 {
					t.Errorf("Row1 Column2 = %v, want 7", values["Column2"])
				}
				if got := rows[0].Cell("Column3").NumberFormat(); got != "$#,##0.00" {
					t.Errorf("Row1 Column3 format = %q, want $#,##0.00", got)
				}
				row2 := rows[1].Cell("Column1").Value()
				if row2 != "Gadget" && row2 != "Widget" {
					t.Errorf("Row2 Column1 = %v, want Gadget or Widget", row2)
				}
			},
		},
		{
			name: "merged_cells.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.MergeCells("A2:B2")
				_ = sheet.Cell("A2").SetValue("Merged")
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if !hasMergedRange(sheet, "A2:B2") {
					t.Error("Expected merged range A2:B2 after round-trip")
				}
			},
		},
		{
			name: "named_ranges.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				wb.AddNamedRange("FixtureRange2", "'Sheet1'!$A$1")
			},
			verify: func(t *testing.T, wb Workbook) {
				if !hasNamedRange(wb, "FixtureRange2") {
					t.Error("Expected named range after round-trip")
				}
			},
		},
		{
			name: "comments.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.Cell("B2").SetComment("Fixture comment", "Tester")
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if comment, ok := sheet.Cell("B2").Comment(); !ok || comment.Text() != "Fixture comment" {
					t.Error("Expected comment after round-trip")
				}
			},
		},
		{
			name: "formulas.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.Cell("A1").SetValue(5)
				_ = sheet.Cell("A2").SetValue(7)
				_ = sheet.Cell("A3").SetFormula("A1+A2")
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if sheet.Cell("A3").Formula() != "A1+A2" {
					t.Errorf("A3 formula = %q, want A1+A2", sheet.Cell("A3").Formula())
				}
			},
		},
		{
			name: "conditional_format.xlsx",
			mutate: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				_ = sheet.Cell("A1").SetValue(10)
			},
			verify: func(t *testing.T, wb Workbook) {
				sheet := wb.SheetsRaw()[0]
				if !hasConditionalFormatting(sheet) {
					t.Error("Expected conditional formatting after round-trip")
				}
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			wb, err := Open(fixtureRoundTripPath(tc.name))
			if err != nil {
				t.Fatalf("Open() error = %v", err)
			}
			if tc.mutate != nil {
				tc.mutate(t, wb)
			}

			outPath := filepath.Join(t.TempDir(), tc.name)
			if err := wb.SaveAs(outPath); err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			if err := wb.Close(); err != nil {
				t.Fatalf("Close() error = %v", err)
			}

			round, err := Open(outPath)
			if err != nil {
				t.Fatalf("Open(roundtrip) error = %v", err)
			}
			defer round.Close()

			if tc.verify != nil {
				tc.verify(t, round)
			}
		})
	}
}

func hasNamedRange(wb Workbook, name string) bool {
	for _, nr := range wb.NamedRanges() {
		if nr.Name() == name {
			return true
		}
	}
	return false
}

func hasMergedRange(sheet Worksheet, ref string) bool {
	for _, rng := range sheet.MergedCells() {
		if rng.Reference() == ref {
			return true
		}
	}
	return false
}

func hasConditionalFormatting(sheet Worksheet) bool {
	ws, ok := sheet.(*worksheetImpl)
	if !ok || ws.worksheet == nil {
		return false
	}
	return ws.worksheet.ConditionalFormatting != nil
}

func cellAlignment(sheet Worksheet, ref string) (Alignment, Alignment) {
	ws, ok := sheet.(*worksheetImpl)
	if !ok || ws.workbook == nil {
		return "", ""
	}
	cell, ok := ws.Cell(ref).(*cellImpl)
	if !ok || cell == nil {
		return "", ""
	}
	styles, ok := ws.workbook.Styles().(*stylesImpl)
	if !ok || styles.stylesheet == nil || styles.stylesheet.CellXfs == nil {
		return "", ""
	}
	if cell.cell.S < 0 || cell.cell.S >= len(styles.stylesheet.CellXfs.Xf) {
		return "", ""
	}
	xf := styles.stylesheet.CellXfs.Xf[cell.cell.S]
	if xf.Alignment == nil {
		return "", ""
	}
	return Alignment(xf.Alignment.Horizontal), Alignment(xf.Alignment.Vertical)
}

func fixtureRoundTripPath(name string) string {
	return filepath.Join("..", "..", "testdata", "excel", name)
}
