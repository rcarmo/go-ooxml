package spreadsheet

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func FuzzSpreadsheetFixtureMutation(f *testing.F) {
	fixtures := []string{
		"comments.xlsx",
		"conditional_format.xlsx",
		"data_types.xlsx",
		"formatting.xlsx",
		"formulas.xlsx",
		"merged_cells.xlsx",
		"minimal.xlsx",
		"multiple_sheets.xlsx",
		"named_ranges.xlsx",
		"single_cell.xlsx",
		"tables.xlsx",
	}

	for _, name := range fixtures {
		data, err := os.ReadFile(fixturePath(name))
		if err != nil {
			continue
		}
		f.Add(data, uint16(0), byte(0))
	}

	f.Fuzz(func(t *testing.T, data []byte, offset uint16, xor byte) {
		if len(data) == 0 || len(data) > 8<<20 {
			return
		}

		mutated := testutil.MutateBytes(data, offset, xor)
		wb, err := OpenReader(bytes.NewReader(mutated), int64(len(mutated)))
		if err != nil {
			return
		}
		defer wb.Close()

		_ = wb.Sheets()
		_ = wb.Tables()

		path := filepath.Join(t.TempDir(), "fixture.xlsx")
		if err := wb.SaveAs(path); err != nil {
			return
		}

		reopen, err := Open(path)
		if err != nil {
			return
		}
		defer reopen.Close()

		_ = reopen.Sheets()
		_ = reopen.Tables()
	})
}

func fixturePath(name string) string {
	return filepath.Join("..", "..", "testdata", "excel", name)
}
