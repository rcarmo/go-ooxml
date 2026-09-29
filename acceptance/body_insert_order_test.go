package acceptance

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

func bodyInsertOrderSteps(sc *godog.ScenarioContext) {
	var doc document.Document
	var body document.Body
	var counts []int
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, body, counts = nil, nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a new Word body with no elements$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		body = doc.Body()
		if body == nil {
			return fmt.Errorf("new Word body missing")
		}
		if got := body.ElementCount(); got != 0 {
			return fmt.Errorf("initial body element count = %d, want 0", got)
		}
		return nil
	})
	sc.Step(`^First and Third paragraphs are appended, then Second is inserted at index one$`, func() error {
		if body == nil {
			return fmt.Errorf("new Word body missing")
		}
		body.AddParagraph().SetText("First")
		counts = append(counts, body.ElementCount())
		body.AddParagraph().SetText("Third")
		counts = append(counts, body.ElementCount())
		inserted := body.InsertParagraphAt(1)
		if inserted == nil {
			return fmt.Errorf("inserted paragraph missing")
		}
		inserted.SetText("Second")
		counts = append(counts, body.ElementCount())
		return nil
	})
	sc.Step(`^element counts after each operation are one, two and three$`, func() error {
		return checkBodyInsertCounts(counts)
	})
	sc.Step(`^paragraph texts in order equal First, Second and Third$`, func() error {
		return checkBodyInsertTexts(body)
	})
}

func checkBodyInsertCounts(counts []int) error {
	want := []int{1, 2, 3}
	if len(counts) != len(want) {
		return fmt.Errorf("body count observations = %d, want 3", len(counts))
	}
	for i, expected := range want {
		if counts[i] != expected {
			return fmt.Errorf("body element count after operation %d = %d, want %d", i+1, counts[i], expected)
		}
	}
	return nil
}

func checkBodyInsertTexts(body document.Body) error {
	if body == nil {
		return fmt.Errorf("new Word body missing")
	}
	want := []string{"First", "Second", "Third"}
	paragraphs := body.Paragraphs()
	if len(paragraphs) != len(want) {
		return fmt.Errorf("body paragraph count = %d, want 3", len(paragraphs))
	}
	for i, expected := range want {
		if paragraphs[i] == nil {
			return fmt.Errorf("body paragraph %d missing", i)
		}
		if got := paragraphs[i].Text(); got != expected {
			return fmt.Errorf("body paragraph %d text = %q, want %q", i, got, expected)
		}
	}
	return nil
}

func guardBodyInsertOrderCase(name string, steps []string) error {
	const expectedName = "Insert a body paragraph between two existing paragraphs"
	if name != expectedName {
		return fmt.Errorf("unexpected body-insert case %q", name)
	}
	return guardParagraphSteps(name, steps, []string{
		"a new Word body with no elements",
		"First and Third paragraphs are appended, then Second is inserted at index one",
		"element counts after each operation are one, two and three",
		"paragraph texts in order equal First, Second and Third",
	})
}

func TestBodyInsertOrderNegativeControls(t *testing.T) {
	doc, err := document.New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	body := doc.Body()
	if body == nil || body.ElementCount() != 0 {
		t.Fatal("expected fresh empty body")
	}
	if err := checkBodyInsertCounts([]int{1, 2, 3}); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name   string
		counts []int
	}{
		{"first", []int{0, 2, 3}}, {"second", []int{1, 3, 3}}, {"third", []int{1, 2, 2}},
		{"missing", []int{1, 2}}, {"extra", []int{1, 2, 3, 4}},
	} {
		if err := checkBodyInsertCounts(tc.counts); err == nil {
			t.Errorf("%s wrong counts passed", tc.name)
		}
	}
	if err := checkBodyInsertTexts(body); err == nil {
		t.Error("empty body passed ordered texts")
	}
	body.AddParagraph().SetText("First")
	body.AddParagraph().SetText("Third")
	body.InsertParagraphAt(1).SetText("Second")
	if err := checkBodyInsertTexts(body); err != nil {
		t.Fatal(err)
	}
	body.Paragraphs()[1].SetText("Wrong")
	if err := checkBodyInsertTexts(body); err == nil || !strings.Contains(err.Error(), "paragraph 1 text") {
		t.Errorf("changed middle text passed: %v", err)
	}
	body.Paragraphs()[1].SetText("Second")
	body.Paragraphs()[0].SetText("Third")
	if err := checkBodyInsertTexts(body); err == nil || !strings.Contains(err.Error(), "paragraph 0 text") {
		t.Errorf("changed first text passed: %v", err)
	}
	body.Paragraphs()[0].SetText("First")
	body.Paragraphs()[2].SetText("First")
	if err := checkBodyInsertTexts(body); err == nil || !strings.Contains(err.Error(), "paragraph 2 text") {
		t.Errorf("changed final text passed: %v", err)
	}
	body.AddParagraph().SetText("Fourth")
	if err := checkBodyInsertTexts(body); err == nil || !strings.Contains(err.Error(), "paragraph count") {
		t.Errorf("extra paragraph passed: %v", err)
	}
	if err := checkBodyInsertTexts(nil); err == nil {
		t.Error("missing body passed")
	}
	base := []string{"a new Word body with no elements", "First and Third paragraphs are appended, then Second is inserted at index one", "element counts after each operation are one, two and three", "paragraph texts in order equal First, Second and Third"}
	if err := guardBodyInsertOrderCase("Insert a body paragraph between two existing paragraphs", base); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name  string
		steps []string
	}{
		{"wrong insert index", []string{base[0], "First and Third paragraphs are appended, then Second is inserted at index zero", base[2], base[3]}},
		{"wrong counts", []string{base[0], base[1], "element counts after each operation are zero, two and three", base[3]}},
		{"wrong order", []string{base[0], base[1], base[2], "paragraph texts in order equal First, Third and Second"}},
		{"missing step", base[:3]},
		{"saved XML step", append(append([]string{}, base...), "the document is saved and reopened")},
	} {
		if err := guardBodyInsertOrderCase("Insert a body paragraph between two existing paragraphs", tc.steps); err == nil {
			t.Errorf("%s passed step guard", tc.name)
		}
	}
	if err := guardBodyInsertOrderCase("Insert a body paragraph after the last existing paragraph", base); err == nil {
		t.Error("changed case name passed")
	}
}
