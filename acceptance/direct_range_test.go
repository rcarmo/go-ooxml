package acceptance

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/formula"
)

// The canonical feature is selected by ID, not by its feature-level @planned tag.
// Keep the row guard independent of the parser's output and of the Go native tests.
func guardDirectRangeCase(id string, p *messages.Pickle, path string) error {
	parse := map[string]struct {
		source, sheet, axis string
		row, col            int
	}{
		"absolute column":     {"$B:$B", "", "column", 0, 2},
		"reversed whole rows": {"$3:$1", "", "row", 3, 0},
		"quoted column":       {"'O''Brien'!$B:$B", "O'Brien", "column", 0, 2},
		"reversed rectangle":  {"A3:B1", "", "cell", 3, 1},
		"sheet absolute cell": {"Sheet!$A$1", "Sheet", "cell", 1, 1},
		"numeric sheet name":  {"'3'!1:1", "3", "row", 1, 0},
	}
	refuse := map[string]string{
		"mixed endpoint": "A:B1", "formula expression": "SUM(A1)",
		"external book": "[Book]Sheet!A1", "sheet span": "Sheet1:Sheet2!A1",
		"spill reference": "A1#", "union": "A1,B2", "invalid whole axis": "XFE:XFE",
	}
	var steps []string
	valid := false
	if id == directRangeParsingCaseID {
		for variant, row := range parse {
			if p.Name == "A "+variant+" direct range yields bounded axis and first coordinate" {
				source, sheet := jsonString(row.source), jsonString(row.sheet)
				steps = []string{"the direct range source is JSON " + source, "the direct-range parser reads the source", "its sheet is JSON " + sheet + " and its axis is " + row.axis, fmt.Sprintf("its first coordinate has row %d and column %d", row.row, row.col)}
				valid = true
			}
		}
	} else if id == directRangeRefusalCaseID {
		for variant, source := range refuse {
			if p.Name == "A "+variant+" input is not a direct range in this profile" {
				steps = []string{"the direct range source is JSON " + jsonString(source), "the direct-range parser reads the source", "it returns an error"}
				valid = true
			}
		}
	}
	if !valid || len(p.Steps) != len(steps) {
		return fmt.Errorf("%s: unexpected direct-range row %s %q", path, id, p.Name)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("%s: direct-range row %s %q step %d drift", path, id, p.Name, i+1)
		}
	}
	return nil
}

func jsonString(s string) string {
	b, _ := json.Marshal(s)
	return string(b)
}

func directRangeSteps(sc *godog.ScenarioContext) {
	var source string
	var parsed formula.StaticRange
	var failure error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, parsed, failure = "", formula.StaticRange{}, nil
		return ctx, nil
	})
	sc.Step(`^the direct range source is JSON (.+)$`, func(encoded string) error {
		if err := json.Unmarshal([]byte(encoded), &source); err != nil {
			return fmt.Errorf("decode direct range: %w", err)
		}
		return nil
	})
	sc.Step(`^the direct-range parser reads the source$`, func() error {
		parsed, failure = formula.ParseRange(source)
		return nil
	})
	sc.Step(`^its sheet is JSON (.+) and its axis is (column|row|cell)$`, func(encoded, axis string) error {
		var sheet string
		if err := json.Unmarshal([]byte(encoded), &sheet); err != nil {
			return err
		}
		gotAxis := "cell"
		if parsed.WholeColumns {
			gotAxis = "column"
		} else if parsed.WholeRows {
			gotAxis = "row"
		}
		if failure != nil || parsed.Sheet != sheet || gotAxis != axis {
			return fmt.Errorf("direct range %q got %+v axis %s error %v, want sheet %q axis %s", source, parsed, gotAxis, failure, sheet, axis)
		}
		return nil
	})
	sc.Step(`^its first coordinate has row (\d+) and column (\d+)$`, func(row, col int) error {
		if failure != nil || parsed.First.Row != row || parsed.First.Column != col {
			return fmt.Errorf("direct range %q first %+v error %v, want row %d column %d", source, parsed.First, failure, row, col)
		}
		return nil
	})
	sc.Step(`^it returns an error$`, func() error {
		if failure == nil {
			return fmt.Errorf("direct range %q accepted as %+v", source, parsed)
		}
		return nil
	})
}
