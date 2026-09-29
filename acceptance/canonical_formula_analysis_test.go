package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"

	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/formula"
)

// Guard canonical source rows independently of the analysed references and native tests.
func guardCanonicalFormulaAnalysisCase(id string, p *messages.Pickle, path string) error {
	counts := map[string]struct {
		source string
		count  int
	}{
		"mixed formula":    {`IF(A1="B2",'O''Brien'!$C$4,SUM(D1:E2))`, 3},
		"only strings":     {`"A1"&"Sheet1!B2"`, 0},
		"numeric exponent": {`LOG10(A1)+1E10`, 1},
		"unary arithmetic": {`=-(A1+B2)^2%`, 2},
		"boolean literal":  {`TRUE`, 0},
		"Unicode sheet":    {`'α sheet'!A1`, 1},
	}
	refusals := map[string]string{
		"dynamic function": `INDIRECT(A1)`, "unresolved name": `Name+1`,
		"out-of-grid column": `XFE1`, "out-of-grid row": `A1048577`,
		"open quoted sheet": `'unclosed!A1`, "unsupported spill": `@A1`,
		"unfinished function": `A1+SUM(2`,
	}
	var steps []string
	matched := false
	switch id {
	case formulaAnalysisCountsCaseID:
		for variant, row := range counts {
			if p.Name == "Count static references in "+variant+" without treating strings as references" {
				steps = []string{"the formula source is JSON " + canonicalFormulaJSONString(row.source), "the static formula analyser reads the source", fmt.Sprintf("it returns %d reference records without error", row.count), "every reference has a nonempty byte span inside the original source"}
				matched = true
			}
		}
	case formulaQuotedSheetFlagsCaseID:
		if p.Name == "Decode a quoted sheet and independent absolute reference flags" {
			steps = []string{"the formula source is JSON " + canonicalFormulaJSONString(`'O''Brien'!$B2:C$4`), "the static formula analyser reads the source", "its one reference spans every byte of the original source", "the sheet is O'Brien with first cell column 2 row 2 and absolute column only", "the last cell is column 3 row 4 with absolute row only"}
			matched = true
		}
	case formulaAnalysisRefusalCaseID:
		for variant, source := range refusals {
			if p.Name == "Refuse unsupported static formula syntax "+variant+" without partial references" {
				steps = []string{"the formula source is JSON " + canonicalFormulaJSONString(source), "the static formula analyser reads the source", "it returns an error and zero reference records"}
				matched = true
			}
		}
	}
	if !matched || len(p.Steps) != len(steps) {
		return fmt.Errorf("%s: unexpected canonical formula row %s %q", path, id, p.Name)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("%s: canonical formula row %s %q step %d drift", path, id, p.Name, i+1)
		}
	}
	return nil
}

// The canonical Gherkin table writes literal '&'; Go's default JSON encoder
// HTML-escapes it. Preserve the table spelling while keeping JSON escaping.
func canonicalFormulaJSONString(source string) string {
	var buf bytes.Buffer
	encoder := json.NewEncoder(&buf)
	encoder.SetEscapeHTML(false)
	_ = encoder.Encode(source)
	return string(bytes.TrimSuffix(buf.Bytes(), []byte{'\n'}))
}

func canonicalFormulaAnalysisSteps(sc *godog.ScenarioContext) {
	var source string
	var refs []formula.Reference
	var failure error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, refs, failure = "", nil, nil
		return ctx, nil
	})
	sc.Step(`^the formula source is JSON ("(?:\\.|[^"\\])*")$`, func(encoded string) error {
		if err := json.Unmarshal([]byte(encoded), &source); err != nil {
			return fmt.Errorf("decode formula source: %w", err)
		}
		return nil
	})
	sc.Step(`^the static formula analyser reads the source$`, func() error {
		refs, failure = formula.Analyze(source)
		return nil
	})
	sc.Step(`^the analysis is (accepted without error|refused with zero references)$`, func(outcome string) error {
		if outcome == "accepted without error" && failure == nil || outcome == "refused with zero references" && failure != nil && len(refs) == 0 {
			return nil
		}
		return fmt.Errorf("formula %q: got %d refs, error %v; wanted %s", source, len(refs), failure, outcome)
	})
	sc.Step(`^it returns (\d+) reference records without error$`, func(count int) error {
		if failure != nil || len(refs) != count {
			return fmt.Errorf("formula %q: got %d refs, error %v, want %d", source, len(refs), failure, count)
		}
		return nil
	})
	sc.Step(`^every reference has a nonempty byte span inside the original source$`, func() error {
		for _, ref := range refs {
			if ref.Start < 0 || ref.End <= ref.Start || ref.End > len(source) {
				return fmt.Errorf("formula %q: invalid byte span [%d:%d]", source, ref.Start, ref.End)
			}
		}
		return nil
	})
	sc.Step(`^its one reference spans every byte of the original source$`, func() error {
		if failure != nil || len(refs) != 1 || refs[0].Start != 0 || refs[0].End != len(source) {
			return fmt.Errorf("formula %q: refs %+v error %v", source, refs, failure)
		}
		return nil
	})
	sc.Step(`^the sheet is O'Brien with first cell column 2 row 2 and absolute column only$`, func() error {
		if len(refs) != 1 || refs[0].Sheet != "O'Brien" || refs[0].First.Column != 2 || refs[0].First.Row != 2 || !refs[0].First.AbsoluteColumn || refs[0].First.AbsoluteRow {
			return fmt.Errorf("formula %q: first reference %+v", source, refs)
		}
		return nil
	})
	sc.Step(`^the last cell is column 3 row 4 with absolute row only$`, func() error {
		if len(refs) != 1 || refs[0].Last.Column != 3 || refs[0].Last.Row != 4 || refs[0].Last.AbsoluteColumn || !refs[0].Last.AbsoluteRow {
			return fmt.Errorf("formula %q: last reference %+v", source, refs)
		}
		return nil
	})
	sc.Step(`^it returns an error and zero reference records$`, func() error {
		if failure == nil || len(refs) != 0 {
			return fmt.Errorf("formula %q: accepted/partial refs %+v error %v", source, refs, failure)
		}
		return nil
	})
}
