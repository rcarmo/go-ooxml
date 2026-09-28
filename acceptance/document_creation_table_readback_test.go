package acceptance

import (
	"context"
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

var tableReadbackTexts = [3][3]string{
	{"Header1", "Header2", "Header3"},
	{"A1", "B1", "C1"},
	{"A2", "B2", "C2"},
}

func documentCreationSteps(sc *godog.ScenarioContext, state *tableTextReadbackState) {
	var body document.Body
	var paragraphs int
	var tables int
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		body, paragraphs, tables = nil, 0, 0
		return ctx, nil
	})
	// The table-value binding supplies the exact new-document step.
	sc.Step(`^its body paragraphs and tables are enumerated$`, func() error {
		if state.doc == nil {
			return fmt.Errorf("no Word document")
		}
		body = state.doc.Body()
		if body == nil {
			return fmt.Errorf("no Word body")
		}
		paragraphs, tables = len(body.Paragraphs()), len(body.Tables())
		return nil
	})
	sc.Step(`^the body is present with zero paragraphs and zero tables$`, func() error {
		return checkEmptyBody(body, paragraphs, tables)
	})
}

func checkEmptyBody(body document.Body, paragraphs, tables int) error {
	if body == nil {
		return fmt.Errorf("no Word body")
	}
	if paragraphs != 0 {
		return fmt.Errorf("body paragraphs = %d, want 0", paragraphs)
	}
	if tables != 0 {
		return fmt.Errorf("body tables = %d, want 0", tables)
	}
	return nil
}

type tableTextReadbackState struct {
	doc   document.Document
	table document.Table
}

func tableTextReadbackSteps(sc *godog.ScenarioContext, state *tableTextReadbackState) {
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		state.doc, state.table = nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if state.doc != nil {
			_ = state.doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^its cells contain Header1, Header2, Header3, A1, B1, C1, A2, B2 and C2 in row order$`, func() error {
		if state.table == nil {
			return fmt.Errorf("no Word table")
		}
		for row := range tableReadbackTexts {
			for col, value := range tableReadbackTexts[row] {
				cell := state.table.Cell(row, col)
				if cell == nil {
					return fmt.Errorf("table Cell(%d,%d) missing", row, col)
				}
				cell.SetText(value)
			}
		}
		return nil
	})
	sc.Step(`^exactly one table is readable$`, func() error {
		return checkReopenedTableCount(state.doc)
	})
	sc.Step(`^all nine cell text getters equal their original row-order values$`, func() error {
		return checkReopenedTableTexts(state.doc)
	})
}

func (state *tableTextReadbackState) saveAndReopen(outputDir string) error {
	if state.doc == nil {
		return fmt.Errorf("no table document")
	}
	path := filepath.Join(outputDir, "table-text.docx")
	if err := state.doc.SaveAs(path); err != nil {
		return err
	}
	if err := state.doc.Close(); err != nil {
		return err
	}
	state.doc = nil
	var err error
	state.doc, err = document.Open(path)
	return err
}

func checkReopenedTableCount(doc document.Document) error {
	if doc == nil {
		return fmt.Errorf("no reopened Word document")
	}
	if got := len(doc.Tables()); got != 1 {
		return fmt.Errorf("reopened table count = %d, want 1", got)
	}
	return nil
}

func checkReopenedTableTexts(doc document.Document) error {
	if err := checkReopenedTableCount(doc); err != nil {
		return err
	}
	table := doc.Tables()[0]
	for row := range tableReadbackTexts {
		for col, want := range tableReadbackTexts[row] {
			cell := table.Cell(row, col)
			if cell == nil {
				return fmt.Errorf("reopened Cell(%d,%d) missing", row, col)
			}
			if got := cell.Text(); got != want {
				return fmt.Errorf("reopened Cell(%d,%d).Text = %q, want %q", row, col, got, want)
			}
		}
	}
	return nil
}

func TestDocumentCreationAndTableReadbackNegativeControls(t *testing.T) {
	if err := checkEmptyBody(nil, 0, 0); err == nil || !strings.Contains(err.Error(), "body") {
		t.Fatalf("missing body passed: %v", err)
	}
	for _, tc := range []struct {
		name   string
		change func(document.Document)
		want   string
	}{
		{"paragraph", func(doc document.Document) { doc.AddParagraph() }, "paragraphs"},
		{"table", func(doc document.Document) { doc.AddTable(1, 1) }, "tables"},
	} {
		t.Run("empty-body/"+tc.name, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			body := doc.Body()
			if err := checkEmptyBody(body, len(body.Paragraphs()), len(body.Tables())); err != nil {
				t.Fatal(err)
			}
			tc.change(doc)
			if err := checkEmptyBody(body, len(body.Paragraphs()), len(body.Tables())); err == nil || !strings.Contains(err.Error(), tc.want) {
				t.Fatalf("changed body passed: %v", err)
			}
		})
	}
	for row := range tableReadbackTexts {
		for col, value := range tableReadbackTexts[row] {
			t.Run(fmt.Sprintf("reopened-cell/%d,%d", row, col), func(t *testing.T) {
				doc, err := document.New()
				if err != nil {
					t.Fatal(err)
				}
				table := doc.AddTable(3, 3)
				for r := range tableReadbackTexts {
					for c, v := range tableReadbackTexts[r] {
						table.Cell(r, c).SetText(v)
					}
				}
				path := filepath.Join(t.TempDir(), "table.docx")
				if err := doc.SaveAs(path); err != nil {
					t.Fatal(err)
				}
				if err := doc.Close(); err != nil {
					t.Fatal(err)
				}
				reopened, err := document.Open(path)
				if err != nil {
					t.Fatal(err)
				}
				defer reopened.Close()
				if err := checkReopenedTableCount(reopened); err != nil {
					t.Fatal(err)
				}
				if err := checkReopenedTableTexts(reopened); err != nil {
					t.Fatal(err)
				}
				reopened.Tables()[0].Cell(row, col).SetText(value + "-changed")
				if err := checkReopenedTableTexts(reopened); err == nil || !strings.Contains(err.Error(), fmt.Sprintf("Cell(%d,%d)", row, col)) {
					t.Fatalf("changed cell passed: %v", err)
				}
			})
		}
	}
	t.Run("reopened-table-count", func(t *testing.T) {
		doc, err := document.New()
		if err != nil {
			t.Fatal(err)
		}
		doc.AddTable(1, 1)
		path := filepath.Join(t.TempDir(), "two-tables.docx")
		if err := doc.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		if err := doc.Close(); err != nil {
			t.Fatal(err)
		}
		reopened, err := document.Open(path)
		if err != nil {
			t.Fatal(err)
		}
		defer reopened.Close()
		if err := checkReopenedTableCount(reopened); err != nil {
			t.Fatal(err)
		}
		reopened.AddTable(1, 1)
		changedPath := filepath.Join(t.TempDir(), "changed.docx")
		if err := reopened.SaveAs(changedPath); err != nil {
			t.Fatal(err)
		}
		changed, err := document.Open(changedPath)
		if err != nil {
			t.Fatal(err)
		}
		defer changed.Close()
		if err := checkReopenedTableCount(changed); err == nil || !strings.Contains(err.Error(), "table count") {
			t.Fatalf("extra reopened table passed: %v", err)
		}
	})
}
