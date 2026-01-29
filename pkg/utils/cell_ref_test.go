package utils

import "testing"

func TestColumnToLetter(t *testing.T) {
	tests := []struct {
		col    int
		expect string
	}{
		{1, "A"},
		{2, "B"},
		{26, "Z"},
		{27, "AA"},
		{28, "AB"},
		{52, "AZ"},
		{53, "BA"},
		{702, "ZZ"},
		{703, "AAA"},
		{0, ""},
		{-1, ""},
	}

	for _, tt := range tests {
		got := ColumnToLetter(tt.col)
		if got != tt.expect {
			t.Errorf("ColumnToLetter(%d) = %q, want %q", tt.col, got, tt.expect)
		}
	}
}

func TestLetterToColumn(t *testing.T) {
	tests := []struct {
		letter string
		expect int
	}{
		{"A", 1},
		{"B", 2},
		{"Z", 26},
		{"AA", 27},
		{"AB", 28},
		{"AZ", 52},
		{"BA", 53},
		{"ZZ", 702},
		{"AAA", 703},
		{"a", 1},
		{"aa", 27},
		{"", 0},
	}

	for _, tt := range tests {
		got := LetterToColumn(tt.letter)
		if got != tt.expect {
			t.Errorf("LetterToColumn(%q) = %d, want %d", tt.letter, got, tt.expect)
		}
	}
}

func TestParseCellRef(t *testing.T) {
	tests := []struct {
		ref     string
		want    CellRef
		wantErr bool
	}{
		{"A1", CellRef{Col: 1, Row: 1}, false},
		{"B2", CellRef{Col: 2, Row: 2}, false},
		{"Z100", CellRef{Col: 26, Row: 100}, false},
		{"AA1", CellRef{Col: 27, Row: 1}, false},
		{"$A$1", CellRef{Col: 1, Row: 1, ColAbs: true, RowAbs: true}, false},
		{"$A1", CellRef{Col: 1, Row: 1, ColAbs: true}, false},
		{"A$1", CellRef{Col: 1, Row: 1, RowAbs: true}, false},
		{"Sheet1!A1", CellRef{Sheet: "Sheet1", Col: 1, Row: 1}, false},
		{"Sheet1!$A$1", CellRef{Sheet: "Sheet1", Col: 1, Row: 1, ColAbs: true, RowAbs: true}, false},
		{"", CellRef{}, true},
		{"A", CellRef{}, true},
		{"1", CellRef{}, true},
		{"A0", CellRef{}, true},
	}

	for _, tt := range tests {
		got, err := ParseCellRef(tt.ref)
		if (err != nil) != tt.wantErr {
			t.Errorf("ParseCellRef(%q) error = %v, wantErr %v", tt.ref, err, tt.wantErr)
			continue
		}
		if !tt.wantErr && got != tt.want {
			t.Errorf("ParseCellRef(%q) = %+v, want %+v", tt.ref, got, tt.want)
		}
	}
}

func TestCellRefString(t *testing.T) {
	tests := []struct {
		ref    CellRef
		expect string
	}{
		{CellRef{Col: 1, Row: 1}, "A1"},
		{CellRef{Col: 27, Row: 100}, "AA100"},
		{CellRef{Col: 1, Row: 1, ColAbs: true}, "$A1"},
		{CellRef{Col: 1, Row: 1, RowAbs: true}, "A$1"},
		{CellRef{Col: 1, Row: 1, ColAbs: true, RowAbs: true}, "$A$1"},
		{CellRef{Sheet: "Sheet1", Col: 1, Row: 1}, "Sheet1!A1"},
	}

	for _, tt := range tests {
		got := tt.ref.String()
		if got != tt.expect {
			t.Errorf("%+v.String() = %q, want %q", tt.ref, got, tt.expect)
		}
	}
}

func TestParseRangeRef(t *testing.T) {
	tests := []struct {
		ref     string
		wantErr bool
	}{
		{"A1:B2", false},
		{"A1:Z100", false},
		{"$A$1:$B$2", false},
		{"Sheet1!A1:B2", false},
		{"A1", true},
		{"", true},
	}

	for _, tt := range tests {
		_, err := ParseRangeRef(tt.ref)
		if (err != nil) != tt.wantErr {
			t.Errorf("ParseRangeRef(%q) error = %v, wantErr %v", tt.ref, err, tt.wantErr)
		}
	}
}

func TestRangeRefContains(t *testing.T) {
	rng, _ := ParseRangeRef("B2:D4")
	
	tests := []struct {
		ref    string
		expect bool
	}{
		{"B2", true},
		{"C3", true},
		{"D4", true},
		{"A1", false},
		{"E5", false},
		{"B5", false},
	}

	for _, tt := range tests {
		cell, _ := ParseCellRef(tt.ref)
		got := rng.Contains(cell)
		if got != tt.expect {
			t.Errorf("range %s.Contains(%s) = %v, want %v", rng.String(), tt.ref, got, tt.expect)
		}
	}
}
