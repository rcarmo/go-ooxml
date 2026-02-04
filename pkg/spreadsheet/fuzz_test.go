package spreadsheet

import (
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzWorksheetCellAccess(f *testing.F) {
	seeds := []string{"A1", "B2", "AA10", "XFD1048576"}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, ref string) {
		w, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer w.Close()

		cell := w.SheetsRaw()[0].Cell(ref)
		if cell == nil {
			return
		}

		parsed, err := utils.ParseCellRef(ref)
		if err != nil {
			t.Fatalf("ParseCellRef(%q) failed after Cell(): %v", ref, err)
		}

		if cell.Row() != parsed.Row || cell.Column() != parsed.Col {
			t.Fatalf("Cell(%q) row/col = (%d,%d), want (%d,%d)", ref, cell.Row(), cell.Column(), parsed.Row, parsed.Col)
		}
	})
}

func FuzzWorksheetRangeAccess(f *testing.F) {
	seeds := []string{"A1:B2", "C3:D4", "AA10:AB11"}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, ref string) {
		w, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer w.Close()

		rng := w.SheetsRaw()[0].Range(ref)
		if rng == nil {
			return
		}

		parsed, err := utils.ParseRangeRef(ref)
		if err != nil {
			t.Fatalf("ParseRangeRef(%q) failed after Range(): %v", ref, err)
		}

		if rng.RowCount() != parsed.RowCount() || rng.ColumnCount() != parsed.ColumnCount() {
			t.Fatalf("Range(%q) size mismatch: got %dx%d want %dx%d", ref, rng.RowCount(), rng.ColumnCount(), parsed.RowCount(), parsed.ColumnCount())
		}
	})
}
