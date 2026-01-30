package utils

import "testing"

func FuzzColumnLetterRoundTrip(f *testing.F) {
	seeds := []int{1, 26, 27, 52, 702, 703, 16384}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, col int) {
		if col < 1 || col > 16384 {
			return
		}

		letters := ColumnToLetter(col)
		if letters == "" {
			t.Fatalf("ColumnToLetter(%d) returned empty string", col)
		}

		got := LetterToColumn(letters)
		if got != col {
			t.Fatalf("LetterToColumn(%q) = %d, want %d", letters, got, col)
		}
	})
}

func FuzzParseCellRef(f *testing.F) {
	seeds := []string{
		"A1",
		"Z99",
		"$A$1",
		"Sheet1!B2",
		"Sheet 1!$C$3",
		"Data!AA10",
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, input string) {
		ref, err := ParseCellRef(input)
		if err != nil {
			return
		}

		if ref.Row < 1 || ref.Col < 1 {
			t.Fatalf("ParseCellRef(%q) produced invalid row/col: %d/%d", input, ref.Row, ref.Col)
		}

		roundTrip := ref.String()
		parsed, err := ParseCellRef(roundTrip)
		if err != nil {
			t.Fatalf("ParseCellRef round-trip failed for %q: %v", roundTrip, err)
		}

		if parsed.Row != ref.Row || parsed.Col != ref.Col || parsed.Sheet != ref.Sheet || parsed.ColAbs != ref.ColAbs || parsed.RowAbs != ref.RowAbs {
			t.Fatalf("ParseCellRef round-trip mismatch: %+v vs %+v", ref, parsed)
		}
	})
}

func FuzzParseRangeRef(f *testing.F) {
	seeds := []string{
		"A1:B2",
		"Sheet1!$A$1:$B$2",
		"C3:D4",
	}
	for _, seed := range seeds {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, input string) {
		rng, err := ParseRangeRef(input)
		if err != nil {
			return
		}

		if rng.Start.Row < 1 || rng.Start.Col < 1 || rng.End.Row < 1 || rng.End.Col < 1 {
			t.Fatalf("ParseRangeRef(%q) produced invalid bounds: %+v", input, rng)
		}

		roundTrip := rng.String()
		parsed, err := ParseRangeRef(roundTrip)
		if err != nil {
			t.Fatalf("ParseRangeRef round-trip failed for %q: %v", roundTrip, err)
		}

		if parsed.Start != rng.Start || parsed.End != rng.End {
			t.Fatalf("ParseRangeRef round-trip mismatch: %+v vs %+v", rng, parsed)
		}
	})
}
