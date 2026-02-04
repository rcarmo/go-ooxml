package spreadsheet

import (
	"path/filepath"
	"testing"
)

func FuzzWorkbookSaveOpen(f *testing.F) {
	f.Add("Sheet1", "A1", "Hello")
	f.Add("Data", "B2", "42")

	f.Fuzz(func(t *testing.T, sheetName, cellRef, value string) {
		if len(sheetName) == 0 || len(sheetName) > 64 || len(cellRef) > 16 || len(value) > 256 {
			return
		}

		wb, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer wb.Close()

		sheet := wb.AddSheet(sheetName)
		cell := sheet.Cell(cellRef)
		if cell != nil {
			_ = cell.SetValue(value)
		}

		path := filepath.Join(t.TempDir(), "fuzz.xlsx")
		if err := wb.SaveAs(path); err != nil {
			t.Fatalf("SaveAs() error = %v", err)
		}

		reopen, err := Open(path)
		if err != nil {
			t.Fatalf("Open() error = %v", err)
		}
		defer reopen.Close()

		if reopen.SheetCount() < 1 {
			t.Fatalf("expected at least one sheet")
		}
	})
}
