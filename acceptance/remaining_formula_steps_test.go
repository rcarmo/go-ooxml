package acceptance

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/formula"
)

func remainingFormulaSteps(sc *godog.ScenarioContext) {
	var source, contextSheet, output string
	var failure error
	var matrix []matrixFormula
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, contextSheet, output = "", "", ""
		failure, matrix = nil, nil
		return ctx, nil
	})
	sc.Step(`^the formula source is JSON (.+) in context sheet Main$`, func(encoded string) error {
		if err := json.Unmarshal([]byte(encoded), &source); err != nil {
			return err
		}
		contextSheet = "Main"
		return nil
	})
	sc.Step(`^the static remapper inserts (row|column|bad) at (\d+) by (\d+) on sheet JSON (.+)$`, func(axis string, at, count int, encoded string) error {
		var sheet string
		if err := json.Unmarshal([]byte(encoded), &sheet); err != nil {
			return err
		}
		output, failure = formula.InsertReferences(source, contextSheet, formula.Insertion{Sheet: sheet, Axis: axis, At: at, Count: count})
		return nil
	})
	sc.Step(`^the complete replacement expression equals JSON (.+)$`, func(encoded string) error {
		var expected string
		if err := json.Unmarshal([]byte(encoded), &expected); err != nil {
			return err
		}
		if failure != nil || output != expected {
			return fmt.Errorf("remap %q: got %q, error %v; want %q", source, output, failure, expected)
		}
		return nil
	})
	sc.Step(`^it returns an error and an empty replacement expression$`, func() error {
		if failure == nil || output != "" {
			return fmt.Errorf("remap %q: got %q, error %v; wanted empty refusal", source, output, failure)
		}
		return nil
	})
	sc.Step(`^cell tokens A1, \$B2, C\$3 and \$XFD\$9$`, func() error { return nil })
	sc.Step(`^optional prefixes empty, Main! and 'Input Data'! with operators \+, -, \*, /, & and >=$`, func() error { return nil })
	sc.Step(`^the static analyser checks all 3 by 4 by 4 by 6 source expressions$`, func() error {
		matrix = makeFormulaMatrix()
		if len(matrix) != 3*4*4*6 {
			return fmt.Errorf("matrix size %d", len(matrix))
		}
		for i := range matrix {
			matrix[i].refs, matrix[i].err = formula.Analyze(matrix[i].source)
			if matrix[i].err != nil {
				return fmt.Errorf("matrix %d %q: %w", i, matrix[i].source, matrix[i].err)
			}
		}
		return nil
	})
	sc.Step(`^every expression has two references whose source slices each parse as one matching reference$`, func() error {
		if len(matrix) != 288 {
			return fmt.Errorf("matrix missing: %d", len(matrix))
		}
		for i, entry := range matrix {
			if entry.err != nil || len(entry.refs) != 2 {
				return fmt.Errorf("matrix %d %q: refs %+v error %v", i, entry.source, entry.refs, entry.err)
			}
			for n, ref := range entry.refs {
				want := entry.expected[n]
				if ref.Start != want.Start || ref.End != want.End || ref.Sheet != want.Sheet || ref.First != want.First || ref.Last != want.Last || entry.source[ref.Start:ref.End] != entry.raw[n] {
					return fmt.Errorf("matrix %d %q ref %d: got %+v, want %+v slice %q", i, entry.source, n, ref, want, entry.raw[n])
				}
				parsed, err := formula.Analyze(entry.raw[n])
				if err != nil || len(parsed) != 1 || parsed[0].Sheet != want.Sheet || parsed[0].First != want.First || parsed[0].Last != want.Last || parsed[0].Start != 0 || parsed[0].End != len(entry.raw[n]) {
					return fmt.Errorf("matrix %d %q reparse %q: %+v %v", i, entry.source, entry.raw[n], parsed, err)
				}
			}
		}
		return nil
	})
	sc.Step(`^inserting one row at 100 on Main leaves each original expression byte-identical$`, func() error {
		if len(matrix) != 288 {
			return fmt.Errorf("matrix missing: %d", len(matrix))
		}
		for i, entry := range matrix {
			got, err := formula.InsertReferences(entry.source, "Main", formula.Insertion{Sheet: "Main", Axis: "row", At: 100, Count: 1})
			if err != nil || got != entry.source {
				return fmt.Errorf("matrix %d no-op: got %q, error %v, want %q", i, got, err, entry.source)
			}
		}
		return nil
	})
	sc.Step(`^wrapping each expression in SUM preserves its references after adjusting their byte spans$`, func() error {
		if len(matrix) != 288 {
			return fmt.Errorf("matrix missing: %d", len(matrix))
		}
		for i, entry := range matrix {
			wrapped, err := formula.Analyze("SUM(" + entry.source + ")")
			if err != nil || len(wrapped) != 2 {
				return fmt.Errorf("matrix %d wrapped: %+v %v", i, wrapped, err)
			}
			for n, ref := range wrapped {
				want := entry.expected[n]
				want.Start += len("SUM(")
				want.End += len("SUM(")
				if ref != want {
					return fmt.Errorf("matrix %d wrapped ref %d: got %+v, want %+v", i, n, ref, want)
				}
			}
		}
		return nil
	})
}

type matrixFormula struct {
	source   string
	raw      [2]string
	expected [2]formula.Reference
	refs     []formula.Reference
	err      error
}

// Construct expected spans and endpoints directly from fixed literal tokens.
// Neither parser output nor cellString is used to derive the oracle.
func makeFormulaMatrix() []matrixFormula {
	cells := []struct {
		text string
		cell formula.Cell
	}{
		{"A1", formula.Cell{Row: 1, Column: 1}},
		{"$B2", formula.Cell{Row: 2, Column: 2, AbsoluteColumn: true}},
		{"C$3", formula.Cell{Row: 3, Column: 3, AbsoluteRow: true}},
		{"$XFD$9", formula.Cell{Row: 9, Column: 16384, AbsoluteColumn: true, AbsoluteRow: true}},
	}
	prefixes := []struct{ text, sheet string }{{"", ""}, {"Main!", "Main"}, {"'Input Data'!", "Input Data"}}
	ops := []string{"+", "-", "*", "/", "&", ">="}
	out := make([]matrixFormula, 0, len(prefixes)*len(cells)*len(cells)*len(ops))
	for _, prefix := range prefixes {
		for _, left := range cells {
			for _, right := range cells {
				for _, op := range ops {
					a := prefix.text + left.text
					b := right.text
					source := strings.Join([]string{a, op, b}, "")
					out = append(out, matrixFormula{source: source, raw: [2]string{a, b}, expected: [2]formula.Reference{
						{Sheet: prefix.sheet, First: left.cell, Last: left.cell, Start: 0, End: len(a)},
						{First: right.cell, Last: right.cell, Start: len(a) + len(op), End: len(source)},
					}})
				}
			}
		}
	}
	return out
}
