package utils

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// CellRef represents a cell reference (e.g., "A1", "Sheet1!B5").
type CellRef struct {
	Sheet  string // Optional sheet name
	Col    int    // 1-based column number
	Row    int    // 1-based row number
	ColAbs bool   // Is column absolute ($A)
	RowAbs bool   // Is row absolute ($1)
}

var cellRefRegex = regexp.MustCompile(`^(?:([^!]+)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)$`)

// ParseCellRef parses "A1", "$A$1", "Sheet1!A1" formats.
func ParseCellRef(ref string) (CellRef, error) {
	matches := cellRefRegex.FindStringSubmatch(ref)
	if matches == nil {
		return CellRef{}, ErrInvalidCellRef
	}

	col := LetterToColumn(matches[3])
	if col == 0 {
		return CellRef{}, ErrInvalidCellRef
	}

	row, err := strconv.Atoi(matches[5])
	if err != nil || row < 1 {
		return CellRef{}, ErrInvalidCellRef
	}

	return CellRef{
		Sheet:  matches[1],
		Col:    col,
		Row:    row,
		ColAbs: matches[2] == "$",
		RowAbs: matches[4] == "$",
	}, nil
}

// String returns the A1-style string representation.
func (c CellRef) String() string {
	var sb strings.Builder

	if c.Sheet != "" {
		sb.WriteString(c.Sheet)
		sb.WriteByte('!')
	}

	if c.ColAbs {
		sb.WriteByte('$')
	}
	sb.WriteString(ColumnToLetter(c.Col))

	if c.RowAbs {
		sb.WriteByte('$')
	}
	sb.WriteString(strconv.Itoa(c.Row))

	return sb.String()
}

// ColumnToLetter converts 1-based column number to letter(s).
// 1 -> A, 26 -> Z, 27 -> AA, etc.
func ColumnToLetter(col int) string {
	if col < 1 {
		return ""
	}

	result := ""
	for col > 0 {
		col-- // Adjust for 0-based calculation
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// LetterToColumn converts letter(s) to 1-based column number.
// A -> 1, Z -> 26, AA -> 27, etc.
func LetterToColumn(letter string) int {
	letter = strings.ToUpper(letter)
	result := 0
	for _, c := range letter {
		if c < 'A' || c > 'Z' {
			return 0
		}
		result = result*26 + int(c-'A'+1)
	}
	return result
}

// RangeRef represents a range reference (e.g., "A1:B5").
type RangeRef struct {
	Start CellRef
	End   CellRef
}

// ParseRangeRef parses "A1:B5" or "Sheet1!A1:B5" format.
func ParseRangeRef(ref string) (RangeRef, error) {
	// Handle sheet prefix
	sheet := ""
	if idx := strings.Index(ref, "!"); idx != -1 {
		sheet = ref[:idx]
		ref = ref[idx+1:]
	}

	parts := strings.Split(ref, ":")
	if len(parts) != 2 {
		return RangeRef{}, ErrInvalidRange
	}

	// Add sheet back for parsing if present
	startRef := parts[0]
	endRef := parts[1]
	if sheet != "" {
		startRef = sheet + "!" + startRef
		endRef = sheet + "!" + endRef
	}

	start, err := ParseCellRef(startRef)
	if err != nil {
		return RangeRef{}, ErrInvalidRange
	}

	end, err := ParseCellRef(endRef)
	if err != nil {
		return RangeRef{}, ErrInvalidRange
	}

	return RangeRef{Start: start, End: end}, nil
}

// String returns the range string representation.
func (r RangeRef) String() string {
	if r.Start.Sheet != "" {
		// Only include sheet once
		startNoSheet := r.Start
		startNoSheet.Sheet = ""
		endNoSheet := r.End
		endNoSheet.Sheet = ""
		return fmt.Sprintf("%s!%s:%s", r.Start.Sheet, startNoSheet.String(), endNoSheet.String())
	}
	return fmt.Sprintf("%s:%s", r.Start.String(), r.End.String())
}

// Contains checks if a cell reference is within the range.
func (r RangeRef) Contains(c CellRef) bool {
	return c.Col >= r.Start.Col && c.Col <= r.End.Col &&
		c.Row >= r.Start.Row && c.Row <= r.End.Row
}

// RowCount returns the number of rows in the range.
func (r RangeRef) RowCount() int {
	return r.End.Row - r.Start.Row + 1
}

// ColumnCount returns the number of columns in the range.
func (r RangeRef) ColumnCount() int {
	return r.End.Col - r.Start.Col + 1
}

// CellRefFromRC creates a cell reference string from row and column (1-based).
func CellRefFromRC(row, col int) string {
	return ColumnToLetter(col) + strconv.Itoa(row)
}
