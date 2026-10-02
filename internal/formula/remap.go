package formula

import (
	"fmt"
	"strconv"
	"strings"
	"unicode"
)

// Insertion describes a structural row/column insertion in one known worksheet.
// One-based At and positive Count are required; references outside that sheet
// retain their exact source spelling. This is not a workbook mutation.
type Insertion struct {
	Sheet     string
	Axis      string
	At, Count int
}

// InsertReferences rewrites only parsed reference coordinates. Absolute markers
// do not prevent structural shifts (they affect copying, not sheet insertion).
// Ranges spanning the insertion expand; strings and foreign sheets stay exact.
func InsertReferences(source, formulaSheet string, change Insertion) (string, error) {
	return insertReferences(source, formulaSheet, change, false)
}

// InsertReferencesUniform applies the stricter uniform lexical profile while
// retaining the legacy entry point's comparison and unchanged endpoint spelling.
func InsertReferencesUniform(source, formulaSheet string, change Insertion) (string, error) {
	return insertReferences(source, formulaSheet, change, true)
}

func simpleLower(s string) string {
	return strings.Map(unicode.ToLower, s)
}

func insertReferences(source, formulaSheet string, change Insertion, uniform bool) (string, error) {
	limit := 1048576
	if change.Axis == "column" {
		limit = 16384
	} else if change.Axis != "row" {
		return "", fmt.Errorf("axis must be row or column")
	}
	if formulaSheet == "" || change.Sheet == "" || change.At < 1 || change.At > limit || change.Count < 1 || change.Count > limit {
		return "", fmt.Errorf("invalid insertion bounds or worksheet")
	}
	var refs []Reference
	var err error
	if uniform {
		refs, err = AnalyzeUniform(source)
	} else {
		refs, err = Analyze(source)
	}
	if err != nil {
		return "", err
	}
	type edit struct {
		lo, hi int
		text   string
	}
	edits := []edit{}
	shift := func(c Cell) (Cell, error) {
		n := c.Row
		if change.Axis == "column" {
			n = c.Column
		}
		if n >= change.At {
			if n > limit-change.Count {
				return Cell{}, fmt.Errorf("reference exceeds worksheet bounds")
			}
			if change.Axis == "row" {
				c.Row += change.Count
			} else {
				c.Column += change.Count
			}
		}
		return c, nil
	}
	for _, ref := range refs {
		sheet := ref.Sheet
		if sheet == "" {
			sheet = formulaSheet
		}
		if uniform {
			if simpleLower(sheet) != simpleLower(change.Sheet) {
				continue
			}
		} else if !strings.EqualFold(sheet, change.Sheet) {
			continue
		}
		first, err := shift(ref.First)
		if err != nil {
			return "", err
		}
		last, err := shift(ref.Last)
		if err != nil {
			return "", err
		}
		if first == ref.First && last == ref.Last {
			continue
		}
		raw := source[ref.Start:ref.End]
		prefix := ""
		coordinates := raw
		// The last ! is the delimiter even when a quoted sheet itself contains !.
		if ref.Sheet != "" {
			at := strings.LastIndexByte(raw, '!')
			if at < 0 {
				return "", fmt.Errorf("missing reference delimiter")
			}
			prefix = raw[:at+1]
			coordinates = raw[at+1:]
		}
		colon := strings.IndexByte(coordinates, ':')
		left, right := coordinates, ""
		if colon >= 0 {
			left, right = coordinates[:colon], coordinates[colon+1:]
		}
		mappedLeft := left
		if uniform || first != ref.First {
			mappedLeft = cellString(first)
		}
		mappedRight := right
		if colon >= 0 && (uniform || last != ref.Last) {
			mappedRight = cellString(last)
		}
		text := prefix + mappedLeft
		if colon >= 0 {
			text += ":" + mappedRight
		}
		edits = append(edits, edit{ref.Start, ref.End, text})
	}
	var out strings.Builder
	cursor := 0
	for _, e := range edits {
		if e.lo < cursor {
			return "", fmt.Errorf("overlapping reference tokens")
		}
		out.WriteString(source[cursor:e.lo])
		out.WriteString(e.text)
		cursor = e.hi
	}
	out.WriteString(source[cursor:])
	return out.String(), nil
}
func cellString(c Cell) string {
	column := c.Column
	var reverse []byte
	for column > 0 {
		column--
		reverse = append(reverse, byte('A'+column%26))
		column /= 26
	}
	for i, j := 0, len(reverse)-1; i < j; i, j = i+1, j-1 {
		reverse[i], reverse[j] = reverse[j], reverse[i]
	}
	var b strings.Builder
	if c.AbsoluteColumn {
		b.WriteByte('$')
	}
	b.Write(reverse)
	if c.AbsoluteRow {
		b.WriteByte('$')
	}
	b.WriteString(strconv.Itoa(c.Row))
	return b.String()
}
