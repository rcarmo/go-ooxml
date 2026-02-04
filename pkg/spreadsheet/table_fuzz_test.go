package spreadsheet

import (
	"path/filepath"
	"testing"
)

func FuzzSpreadsheetTableRoundTrip(f *testing.F) {
	f.Add("A1:C3", "Table1")
	f.Add("B2:D4", "Sales")

	f.Fuzz(func(t *testing.T, ref, name string) {
		if len(ref) == 0 || len(ref) > 32 || len(name) > 64 {
			return
		}

		wb, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer wb.Close()

		sheet, err := wb.Sheet(0)
		if err != nil {
			t.Fatalf("Sheet(0) error = %v", err)
		}

		table := sheet.AddTable(ref, name)
		if table == nil {
			return
		}
		_ = table.AddRow(map[string]interface{}{})

		path := filepath.Join(t.TempDir(), "fuzz.xlsx")
		if err := wb.SaveAs(path); err != nil {
			t.Fatalf("SaveAs() error = %v", err)
		}

		reopen, err := Open(path)
		if err != nil {
			t.Fatalf("Open() error = %v", err)
		}
		defer reopen.Close()

		if len(reopen.Tables()) == 0 {
			t.Fatalf("expected tables after round-trip")
		}
	})
}
