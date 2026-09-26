package formula

import (
	"fmt"
	"strings"
)

// StaticRange is a direct cell/rectangle or a whole-row/whole-column range.
// Missing coordinates in whole-axis ranges are zero; callers supply proven bounds.
// It is separate from expression analysis, which still refuses whole-axis syntax.
type StaticRange struct {
	Sheet                   string
	First, Last             Cell
	WholeRows, WholeColumns bool
}

func ParseRange(source string) (StaticRange, error) {
	fail := func() (StaticRange, error) {
		return StaticRange{}, fmt.Errorf("expected one direct static cell/row/column range")
	}
	if len(source) == 0 || len(source) > 1<<20 {
		return fail()
	}
	tokens, err := lex(source)
	if err != nil {
		return StaticRange{}, err
	}
	tokens = tokens[:len(tokens)-1]
	out := StaticRange{}
	if len(tokens) >= 2 && tokens[1].kind == "punct" && tokens[1].text == "!" {
		sheet := tokens[0]
		if (sheet.kind != "word" && sheet.kind != "sheet") || sheet.text == "" || strings.ContainsAny(sheet.text, "[]:*?/\\") {
			return fail()
		}
		out.Sheet = sheet.text
		tokens = tokens[2:]
	}
	if len(tokens) != 1 && len(tokens) != 3 {
		return fail()
	}
	if len(tokens) == 3 && (tokens[1].kind != "punct" || tokens[1].text != ":") {
		return fail()
	}
	first, axis, err := rangeEndpoint(tokens[0])
	if err != nil {
		return fail()
	}
	last := first
	if len(tokens) == 1 {
		if axis != "cell" {
			return fail()
		}
	} else {
		var lastAxis string
		last, lastAxis, err = rangeEndpoint(tokens[2])
		if err != nil || axis != lastAxis {
			return fail()
		}
	}
	out.First, out.Last = first, last
	out.WholeColumns = axis == "column"
	out.WholeRows = axis == "row"
	return out, nil
}
func rangeEndpoint(t token) (Cell, string, error) {
	if t.kind != "word" && t.kind != "number" {
		return Cell{}, "", fmt.Errorf("invalid range endpoint")
	}
	if c, err := ParseCell(t.text); err == nil {
		return c, "cell", nil
	}
	raw := strings.TrimPrefix(t.text, "$")
	isLetters, isDigits := raw != "", raw != ""
	for _, r := range raw {
		if !(r >= 'a' && r <= 'z' || r >= 'A' && r <= 'Z') {
			isLetters = false
		}
		if r < '0' || r > '9' {
			isDigits = false
		}
	}
	if isLetters {
		c, err := ParseCell(t.text + "1")
		c.Row = 0
		return c, "column", err
	}
	if isDigits {
		c, err := ParseCell("A" + t.text)
		c.Column = 0
		return c, "row", err
	}
	return Cell{}, "", fmt.Errorf("invalid range endpoint")
}
