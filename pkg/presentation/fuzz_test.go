package presentation

import (
	"path/filepath"
	"testing"
)

func FuzzPresentationRoundTrip(f *testing.F) {
	f.Add("Hello", 1)
	f.Add("Slide", 3)

	f.Fuzz(func(t *testing.T, text string, slides int) {
		if slides < 1 || slides > 5 {
			return
		}
		if len(text) > 1024 {
			return
		}

		pres, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer pres.Close()

		for i := 0; i < slides; i++ {
			slide := pres.AddSlide(0)
			slide.AddTextBox(0, 0, 2000000, 500000).SetText(text)
		}

		path := filepath.Join(t.TempDir(), "fuzz.pptx")
		if err := pres.SaveAs(path); err != nil {
			t.Fatalf("SaveAs() error = %v", err)
		}

		reopen, err := Open(path)
		if err != nil {
			t.Fatalf("Open() error = %v", err)
		}
		defer reopen.Close()

		if reopen.SlideCount() != slides {
			t.Fatalf("SlideCount() = %d, want %d", reopen.SlideCount(), slides)
		}
	})
}

func FuzzPresentationTableOps(f *testing.F) {
	f.Add(2, 2, "Cell")
	f.Add(3, 1, "Value")

	f.Fuzz(func(t *testing.T, rows, cols int, text string) {
		if rows < 1 || rows > 8 || cols < 1 || cols > 8 {
			return
		}
		if len(text) > 512 {
			return
		}

		pres, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer pres.Close()

		slide := pres.AddSlide(0)
		table := slide.AddTable(rows, cols, 0, 0, 4000000, 2000000)
		if table == nil {
			t.Fatalf("AddTable returned nil")
		}
		table.Cell(0, 0).SetText(text)
		if table.RowCount() != rows || table.ColumnCount() != cols {
			t.Fatalf("table size mismatch")
		}
	})
}
