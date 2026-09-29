package acceptance

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

// The selected paragraph scenarios use a fresh in-memory document per case.
func paragraphTextGetterSteps(sc *godog.ScenarioContext) {
	var doc document.Document
	var paragraph document.Paragraph
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, paragraph = nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a new Word paragraph$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		paragraph = doc.AddParagraph()
		if paragraph == nil {
			return fmt.Errorf("new Word paragraph missing")
		}
		return nil
	})
	sc.Step(`^its text is set to JSON (.+)$`, func(raw string) error {
		if paragraph == nil {
			return fmt.Errorf("new Word paragraph missing")
		}
		value, err := decodeParagraphTextJSON(raw)
		if err != nil {
			return err
		}
		paragraph.SetText(value)
		return nil
	})
	sc.Step(`^the paragraph text getter equals JSON (.+)$`, func(raw string) error {
		if paragraph == nil {
			return fmt.Errorf("new Word paragraph missing")
		}
		want, err := decodeParagraphTextJSON(raw)
		if err != nil {
			return err
		}
		return checkParagraphText(paragraph, want)
	})
	paragraphValueSteps(sc, func() document.Paragraph { return paragraph })
}

func decodeParagraphTextJSON(raw string) (string, error) {
	var value *string
	if err := json.Unmarshal([]byte(raw), &value); err != nil {
		return "", fmt.Errorf("paragraph text JSON %q: %w", raw, err)
	}
	if value == nil {
		return "", fmt.Errorf("paragraph text JSON %q: expected string", raw)
	}
	return *value, nil
}

func checkParagraphText(paragraph document.Paragraph, want string) error {
	if paragraph == nil {
		return fmt.Errorf("new Word paragraph missing")
	}
	if got := paragraph.Text(); got != want {
		return fmt.Errorf("paragraph text getter = %q, want %q", got, want)
	}
	return nil
}

// Guard each selected expanded row's name and every step, independently of sibling rows.
func guardParagraphTextRow(name string, steps []string, seen map[string]bool) error {
	variant := strings.TrimSuffix(strings.TrimPrefix(name, "A paragraph reads back "), " text in memory")
	values := map[string]string{"empty": `""`, "simple": `"Hello World"`, "spaced": `"  spaces  "`, "Japanese": `"日本語テキスト"`, "punctuation": `"a < b > c & d"`}
	value, ok := values[variant]
	want := []string{"a new Word paragraph", "its text is set to JSON " + value, "the paragraph text getter equals JSON " + value}
	if !ok || seen[variant] || name != "A paragraph reads back "+variant+" text in memory" || len(steps) != len(want) {
		return fmt.Errorf("unexpected paragraph text row %q", name)
	}
	for i, step := range want {
		if steps[i] != step {
			return fmt.Errorf("unexpected paragraph text step %d for %s", i+1, variant)
		}
	}
	seen[variant] = true
	return nil
}

func TestParagraphTextGetterNegativeControls(t *testing.T) {
	for _, tc := range []struct{ name, text string }{
		{"empty", ""}, {"simple", "Hello World"}, {"spaced", "  spaces  "},
		{"Japanese", "日本語テキスト"}, {"punctuation", "a < b > c & d"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			paragraph := doc.AddParagraph()
			paragraph.SetText(tc.text)
			if err := checkParagraphText(paragraph, tc.text); err != nil {
				t.Fatal(err)
			}
			paragraph.SetText(tc.text + "!")
			if err := checkParagraphText(paragraph, tc.text); err == nil || !strings.Contains(err.Error(), "paragraph text getter") {
				t.Fatalf("altered %s getter passed: %v", tc.name, err)
			}
		})
	}
	if err := checkParagraphText(nil, ""); err == nil {
		t.Fatal("missing paragraph passed")
	}
	base := []string{"a new Word paragraph", `its text is set to JSON "  spaces  "`, `the paragraph text getter equals JSON "  spaces  "`}
	for _, tc := range []struct {
		name  string
		steps []string
	}{
		{"changed setter", []string{base[0], `its text is set to JSON " spaces  "`, base[2]}},
		{"changed getter", []string{base[0], base[1], `the paragraph text getter equals JSON " spaces  "`}},
		{"extra step", append(append([]string{}, base...), "the document is saved and reopened")},
		{"missing step", base[:2]},
	} {
		if err := guardParagraphTextRow("A paragraph reads back spaced text in memory", tc.steps, map[string]bool{}); err == nil {
			t.Errorf("%s passed exact-step guard", tc.name)
		}
	}
	if err := guardParagraphTextRow("A paragraph reads back spaced text in memory", base, map[string]bool{"spaced": true}); err == nil {
		t.Error("duplicate spaced row passed exact-row guard")
	}
	if err := guardParagraphTextRow("A paragraph reads back tabs text in memory", base, map[string]bool{}); err == nil {
		t.Error("unselected native tabs row passed exact-row guard")
	}
	for _, raw := range []string{`null`, `42`, `"unterminated`, `"ok" "extra"`} {
		if _, err := decodeParagraphTextJSON(raw); err == nil {
			t.Errorf("non-string or malformed JSON %q passed", raw)
		}
	}
}
