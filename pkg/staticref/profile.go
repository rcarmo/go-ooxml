// Package staticref exposes the bounded, workbook-free static A1 profile.
// It does not evaluate formulae or edit a workbook.
package staticref

import (
	"errors"
	"fmt"
	"strings"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/formula"
)

type Cell struct {
	Row, Column                 int
	RowAbsolute, ColumnAbsolute bool
}
type Reference struct {
	Sheet, Axis string
	First, Last Cell
	Start, End  int
}
type Range struct {
	Sheet, Axis string
	First, Last Cell
}
type Insertion struct {
	Sheet, Axis string
	At, Count   int
}

type Refusal struct {
	Category string
	Cause    error
}

func (e *Refusal) Error() string            { return e.Cause.Error() }
func (e *Refusal) Unwrap() error            { return e.Cause }
func refuse(category string, e error) error { return &Refusal{Category: category, Cause: e} }

func cell(c formula.Cell) Cell {
	return Cell{Row: c.Row, Column: c.Column, RowAbsolute: c.AbsoluteRow, ColumnAbsolute: c.AbsoluteColumn}
}
func checkInput(source string) error {
	if !utf8.ValidString(source) {
		return refuse("unsupported-static-reference", fmt.Errorf("invalid UTF-8 formula"))
	}
	if len(source) > 1048576 {
		return refuse("static-reference-limit", fmt.Errorf("source UTF-8 byte budget"))
	}
	return nil
}

// Analyze preserves exact UTF-8 byte spans. The production parser owns the
// grammar; this public profile restricts its legacy function surface.
func Analyze(source string) ([]Reference, error) {
	if err := checkInput(source); err != nil {
		return nil, err
	}
	if err := checkFunctions(source); err != nil {
		return nil, err
	}
	refs, err := formula.AnalyzeUniform(source)
	if err != nil {
		return nil, formulaError(err)
	}
	if len(refs) > 10000 {
		return nil, refuse("static-reference-limit", fmt.Errorf("reference budget exceeded"))
	}
	for _, r := range refs {
		if spacedReference(source[r.Start:r.End]) {
			return nil, refuse("unsupported-static-reference", fmt.Errorf("noncontiguous reference"))
		}
	}
	out := make([]Reference, len(refs))
	for i, r := range refs {
		out[i] = Reference{Sheet: r.Sheet, Axis: "cell", First: cell(r.First), Last: cell(r.Last), Start: r.Start, End: r.End}
	}
	return out, nil
}
func formulaError(err error) error {
	var limit *formula.LimitError
	if errors.As(err, &limit) {
		return refuse("static-reference-limit", err)
	}
	return refuse("unsupported-static-reference", err)
}

// ParseRange preserves first/last ordering, absolute flags, and absent axes.
func ParseRange(source string) (Range, error) {
	if !utf8.ValidString(source) {
		return Range{}, refuse("unsupported-direct-range", fmt.Errorf("invalid UTF-8 direct range"))
	}
	if len(source) > 1048576 {
		return Range{}, refuse("unsupported-direct-range", fmt.Errorf("source UTF-8 byte budget"))
	}
	if spacedReference(source) || strings.TrimSpace(source) != source || invalidDirectCharacters(source) {
		return Range{}, refuse("unsupported-direct-range", fmt.Errorf("noncontiguous direct range"))
	}
	r, err := formula.ParseRangeUniform(source)
	if err != nil {
		return Range{}, refuse("unsupported-direct-range", err)
	}
	axis := "cell"
	if r.WholeRows {
		axis = "row"
	} else if r.WholeColumns {
		axis = "column"
	}
	return Range{Sheet: r.Sheet, Axis: axis, First: cell(r.First), Last: cell(r.Last)}, nil
}
func InsertReferences(source, context string, change Insertion) (string, error) {
	if err := checkInput(source); err != nil {
		return "", err
	}
	if change.Axis != "row" && change.Axis != "column" || !validSheet(context) || !validSheet(change.Sheet) || change.At < 1 || change.Count < 1 {
		return "", refuse("invalid-reference-insertion", fmt.Errorf("invalid insertion"))
	}
	limit := 1048576
	if change.Axis == "column" {
		limit = 16384
	}
	if change.At > limit || change.Count > limit {
		return "", refuse("invalid-reference-insertion", fmt.Errorf("insertion out of grid"))
	}
	if _, err := Analyze(source); err != nil {
		return "", err
	}
	out, err := formula.InsertReferencesUniform(source, context, formula.Insertion{Sheet: change.Sheet, Axis: change.Axis, At: change.At, Count: change.Count})
	if err != nil {
		return "", refuse("invalid-reference-insertion", err)
	}
	if len(out) > 1048576 {
		return "", refuse("static-reference-limit", fmt.Errorf("output UTF-8 byte budget"))
	}
	return out, nil
}

// Quotes in sheet names permit spaces; no other reference-token space is legal.
func spacedReference(source string) bool {
	quoted := false
	for at := 0; at < len(source); at++ {
		switch source[at] {
		case '\'':
			if quoted && at+1 < len(source) && source[at+1] == '\'' {
				at++
			} else {
				quoted = !quoted
			}
		case ' ', '\t', '\n', '\r':
			if !quoted {
				return true
			}
		}
		if !quoted && source[at] >= 0x7f {
			r, n := utf8.DecodeRuneInString(source[at:])
			if unicode.IsSpace(r) || forbiddenControl(r) {
				return true
			}
			at += n - 1
		}
	}
	return false
}

func invalidDirectCharacters(s string) bool {
	quoted := false
	for at := 0; at < len(s); {
		r, n := utf8.DecodeRuneInString(s[at:])
		if forbiddenControl(r) {
			return true
		}
		if r == '\'' {
			if quoted && at+n < len(s) && s[at+n] == '\'' {
				at += n + 1
				continue
			}
			quoted = !quoted
		} else if !quoted && unicode.IsSpace(r) && r != ' ' {
			return true
		}
		at += n
	}
	return false
}

func validSheet(s string) bool {
	if !utf8.ValidString(s) || strings.TrimSpace(s) == "" || len(utf16.Encode([]rune(s))) > 255 {
		return false
	}
	for _, r := range s {
		if forbiddenControl(r) || strings.ContainsRune("[]:*?/\\", r) {
			return false
		}
	}
	return true
}
