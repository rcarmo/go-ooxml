package acceptance

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

// This case checks direct cell properties in memory, not physical merge geometry.
func tableMergeGetterSteps(sc *godog.ScenarioContext) {
	var doc document.Document
	var table document.Table
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, table = nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a new Word table with three rows and four columns$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		table = doc.AddTable(3, 4)
		return nil
	})
	sc.Step(`^cell zero-zero gets GridSpan three and first-column rows one and two get restart and continue$`, func() error {
		span, restart, continuation, err := mergeCells(table)
		if err != nil {
			return err
		}
		span.SetGridSpan(3)
		restart.SetVerticalMerge(document.VerticalMerge("restart"))
		continuation.SetVerticalMerge(document.VerticalMerge("continue"))
		return nil
	})
	sc.Step(`^GridSpan at zero-zero equals three$`, func() error {
		span, _, _, err := mergeCells(table)
		if err != nil {
			return err
		}
		return checkGridSpan(span, 3)
	})
	sc.Step(`^VerticalMerge at row one is restart and at row two is continue$`, func() error {
		_, restart, continuation, err := mergeCells(table)
		if err != nil {
			return err
		}
		return checkVerticalMerges(restart, continuation)
	})
}

func mergeCells(table document.Table) (document.Cell, document.Cell, document.Cell, error) {
	if table == nil {
		return nil, nil, nil, fmt.Errorf("Word table missing")
	}
	span, restart, continuation := table.Cell(0, 0), table.Cell(1, 0), table.Cell(2, 0)
	if span == nil || restart == nil || continuation == nil {
		return nil, nil, nil, fmt.Errorf("first-column merge cells missing")
	}
	return span, restart, continuation, nil
}

func checkGridSpan(cell document.Cell, want int) error {
	if cell == nil {
		return fmt.Errorf("span cell missing")
	}
	if got := cell.GridSpan(); got != want {
		return fmt.Errorf("GridSpan getter = %d, want %d", got, want)
	}
	return nil
}

func checkVerticalMerges(restart, continuation document.Cell) error {
	if restart == nil || continuation == nil {
		return fmt.Errorf("vertical merge cells missing")
	}
	if got := restart.VerticalMerge(); got != document.VerticalMerge("restart") {
		return fmt.Errorf("row one VerticalMerge getter = %q, want restart", got)
	}
	if got := continuation.VerticalMerge(); got != document.VerticalMerge("continue") {
		return fmt.Errorf("row two VerticalMerge getter = %q, want continue", got)
	}
	return nil
}

func TestTableMergeGetterNegativeControls(t *testing.T) {
	doc, err := document.New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	table := doc.AddTable(3, 4)
	span, restart, continuation, err := mergeCells(table)
	if err != nil {
		t.Fatal(err)
	}
	span.SetGridSpan(3)
	restart.SetVerticalMerge(document.VerticalMerge("restart"))
	continuation.SetVerticalMerge(document.VerticalMerge("continue"))
	if err := checkGridSpan(span, 3); err != nil {
		t.Fatalf("span positive control: %v", err)
	}
	if err := checkVerticalMerges(restart, continuation); err != nil {
		t.Fatalf("vertical merge positive control: %v", err)
	}
	span.SetGridSpan(1)
	if err := checkGridSpan(span, 3); err == nil || !strings.Contains(err.Error(), "GridSpan") {
		t.Fatalf("cleared span passed: %v", err)
	}
	span.SetGridSpan(2)
	if err := checkGridSpan(span, 3); err == nil || !strings.Contains(err.Error(), "GridSpan") {
		t.Fatalf("wrong span passed: %v", err)
	}
	// Restore each changed cell before altering the other, to test both predicates.
	restart.SetVerticalMerge("")
	if err := checkVerticalMerges(restart, continuation); err == nil || !strings.Contains(err.Error(), "row one") {
		t.Fatalf("cleared restart passed: %v", err)
	}
	restart.SetVerticalMerge("restart")
	continuation.SetVerticalMerge("")
	if err := checkVerticalMerges(restart, continuation); err == nil || !strings.Contains(err.Error(), "row two") {
		t.Fatalf("cleared continuation passed: %v", err)
	}
	continuation.SetVerticalMerge("restart")
	if err := checkVerticalMerges(restart, continuation); err == nil || !strings.Contains(err.Error(), "row two") {
		t.Fatalf("wrong continuation passed: %v", err)
	}
}
