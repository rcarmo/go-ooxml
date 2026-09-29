package acceptance

import (
	"fmt"
	"strconv"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

// Only the exact selected in-memory value steps are registered here.
func paragraphValueSteps(sc *godog.ScenarioContext, current func() document.Paragraph) {
	withParagraph := func() (document.Paragraph, error) {
		p := current()
		if p == nil {
			return nil, fmt.Errorf("new Word paragraph missing")
		}
		return p, nil
	}
	sc.Step(`^its alignment is set to (left|center|right|both)$`, func(value string) error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		p.SetAlignment(value)
		return nil
	})
	sc.Step(`^its alignment getter equals (left|center|right|both)$`, func(want string) error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		return checkParagraphAlignment(p, want)
	})
	sc.Step(`^spacing before is set to (0|120|240) and after to (0|120|240)$`, func(before, after string) error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		b, err := strconv.ParseInt(before, 10, 64)
		if err != nil {
			return err
		}
		a, err := strconv.ParseInt(after, 10, 64)
		if err != nil {
			return err
		}
		p.SetSpacingBefore(b)
		p.SetSpacingAfter(a)
		return nil
	})
	sc.Step(`^its before and after getters equal (0|120|240) and (0|120|240)$`, func(before, after string) error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		b, err := strconv.ParseInt(before, 10, 64)
		if err != nil {
			return err
		}
		a, err := strconv.ParseInt(after, 10, 64)
		if err != nil {
			return err
		}
		return checkParagraphSpacing(p, b, a)
	})
	sc.Step(`^KeepLines, PageBreakBefore and WidowControl are set true$`, func() error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		p.SetKeepLines(true)
		p.SetPageBreakBefore(true)
		p.SetWidowControl(true)
		return nil
	})
	sc.Step(`^all three getters are true in memory$`, func() error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		return checkParagraphToggles(p)
	})
	sc.Step(`^runs containing Hello-space, World and exclamation are appended in order$`, func() error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		for _, text := range []string{"Hello ", "World", "!"} {
			p.AddRun().SetText(text)
		}
		return nil
	})
	sc.Step(`^the paragraph has three runs and its text equals Hello World!$`, func() error {
		p, err := withParagraph()
		if err != nil {
			return err
		}
		return checkParagraphRuns(p)
	})
}

func checkParagraphAlignment(p document.Paragraph, want string) error {
	if p == nil {
		return fmt.Errorf("new Word paragraph missing")
	}
	if got := p.Alignment(); got != want {
		return fmt.Errorf("Alignment = %q, want %q", got, want)
	}
	return nil
}
func checkParagraphSpacing(p document.Paragraph, before, after int64) error {
	if p == nil {
		return fmt.Errorf("new Word paragraph missing")
	}
	if got := p.SpacingBefore(); got != before {
		return fmt.Errorf("SpacingBefore = %d, want %d", got, before)
	}
	if got := p.SpacingAfter(); got != after {
		return fmt.Errorf("SpacingAfter = %d, want %d", got, after)
	}
	return nil
}
func checkParagraphToggles(p document.Paragraph) error {
	if p == nil {
		return fmt.Errorf("new Word paragraph missing")
	}
	if !p.KeepLines() {
		return fmt.Errorf("KeepLines getter is false")
	}
	if !p.PageBreakBefore() {
		return fmt.Errorf("PageBreakBefore getter is false")
	}
	if !p.WidowControl() {
		return fmt.Errorf("WidowControl getter is false")
	}
	return nil
}
func checkParagraphRuns(p document.Paragraph) error {
	if p == nil {
		return fmt.Errorf("new Word paragraph missing")
	}
	if got := len(p.Runs()); got != 3 {
		return fmt.Errorf("paragraph run count = %d, want 3", got)
	}
	return checkParagraphText(p, "Hello World!")
}

func guardParagraphSteps(name string, steps, want []string) error {
	if len(steps) != len(want) {
		return fmt.Errorf("unexpected steps for %q: %d, want %d", name, len(steps), len(want))
	}
	for i := range want {
		if steps[i] != want[i] {
			return fmt.Errorf("unexpected step %d for %q", i+1, name)
		}
	}
	return nil
}
func guardParagraphAlignmentRow(name string, steps []string, seen map[string]bool) error {
	values := map[string]string{"left": "left", "center": "center", "right": "right", "justify": "both"}
	for variant, value := range values {
		if name != "A "+variant+" alignment reads back "+value+" in memory" {
			continue
		}
		if seen[variant] {
			return fmt.Errorf("duplicate alignment row %q", name)
		}
		want := []string{"a new Word paragraph", "its alignment is set to " + value, "its alignment getter equals " + value}
		if err := guardParagraphSteps(name, steps, want); err != nil {
			return err
		}
		seen[variant] = true
		return nil
	}
	return fmt.Errorf("unexpected alignment row %q", name)
}
func guardParagraphSpacingRow(name string, steps []string, seen map[string]bool) error {
	values := map[string][2]string{"zero": {"0", "0"}, "six-point": {"120", "120"}, "twelve": {"240", "240"}, "asymmetric": {"240", "120"}}
	for variant, pair := range values {
		if name != "A "+variant+" paragraph reads back spacing in twips" {
			continue
		}
		if seen[variant] {
			return fmt.Errorf("duplicate spacing row %q", name)
		}
		want := []string{"a new Word paragraph", "spacing before is set to " + pair[0] + " and after to " + pair[1], "its before and after getters equal " + pair[0] + " and " + pair[1]}
		if err := guardParagraphSteps(name, steps, want); err != nil {
			return err
		}
		seen[variant] = true
		return nil
	}
	return fmt.Errorf("unexpected spacing row %q", name)
}
func guardParagraphSingleCase(id, name string, steps []string) error {
	var expectedName string
	var want []string
	switch id {
	case paragraphTogglesCaseID:
		expectedName = "Three paragraph flags read true after being enabled"
		want = []string{"a new Word paragraph", "KeepLines, PageBreakBefore and WidowControl are set true", "all three getters are true in memory"}
	case paragraphRunsCaseID:
		expectedName = "Three added runs concatenate in paragraph text"
		want = []string{"a new Word paragraph", "runs containing Hello-space, World and exclamation are appended in order", "the paragraph has three runs and its text equals Hello World!"}
	default:
		return fmt.Errorf("unselected paragraph ID %q", id)
	}
	if name != expectedName {
		return fmt.Errorf("unexpected paragraph case %q", name)
	}
	return guardParagraphSteps(name, steps, want)
}

func TestParagraphValueNegativeControls(t *testing.T) {
	for _, value := range []string{"left", "center", "right", "both"} {
		doc, err := document.New()
		if err != nil {
			t.Fatal(err)
		}
		p := doc.AddParagraph()
		p.SetAlignment(value)
		if err := checkParagraphAlignment(p, value); err != nil {
			t.Fatal(err)
		}
		p.SetAlignment(value + "-changed")
		if err := checkParagraphAlignment(p, value); err == nil {
			t.Errorf("changed alignment %q passed", value)
		}
		_ = doc.Close()
	}
	for _, pair := range [][2]int64{{0, 0}, {120, 120}, {240, 240}, {240, 120}} {
		doc, err := document.New()
		if err != nil {
			t.Fatal(err)
		}
		p := doc.AddParagraph()
		p.SetSpacingBefore(pair[0])
		p.SetSpacingAfter(pair[1])
		if err := checkParagraphSpacing(p, pair[0], pair[1]); err != nil {
			t.Fatal(err)
		}
		p.SetSpacingBefore(pair[0] + 1)
		if err := checkParagraphSpacing(p, pair[0], pair[1]); err == nil || !strings.Contains(err.Error(), "SpacingBefore") {
			t.Errorf("changed before passed: %v", err)
		}
		p.SetSpacingBefore(pair[0])
		p.SetSpacingAfter(pair[1] + 1)
		if err := checkParagraphSpacing(p, pair[0], pair[1]); err == nil || !strings.Contains(err.Error(), "SpacingAfter") {
			t.Errorf("changed after passed: %v", err)
		}
		_ = doc.Close()
	}
	for _, clear := range []struct {
		name  string
		apply func(document.Paragraph)
	}{
		{"KeepLines", func(p document.Paragraph) { p.SetKeepLines(false) }},
		{"PageBreakBefore", func(p document.Paragraph) { p.SetPageBreakBefore(false) }},
		{"WidowControl", func(p document.Paragraph) { p.SetWidowControl(false) }},
	} {
		doc, err := document.New()
		if err != nil {
			t.Fatal(err)
		}
		p := doc.AddParagraph()
		p.SetKeepLines(true)
		p.SetPageBreakBefore(true)
		p.SetWidowControl(true)
		if err := checkParagraphToggles(p); err != nil {
			t.Fatal(err)
		}
		clear.apply(p)
		if err := checkParagraphToggles(p); err == nil || !strings.Contains(err.Error(), clear.name) {
			t.Errorf("cleared %s passed: %v", clear.name, err)
		}
		_ = doc.Close()
	}
	doc, err := document.New()
	if err != nil {
		t.Fatal(err)
	}
	p := doc.AddParagraph()
	for _, text := range []string{"Hello ", "World", "!"} {
		p.AddRun().SetText(text)
	}
	if err := checkParagraphRuns(p); err != nil {
		t.Fatal(err)
	}
	p.AddRun().SetText("")
	if err := checkParagraphRuns(p); err == nil || !strings.Contains(err.Error(), "run count") {
		t.Errorf("four runs passed: %v", err)
	}
	_ = doc.Close()
	doc, err = document.New()
	if err != nil {
		t.Fatal(err)
	}
	p = doc.AddParagraph()
	for _, text := range []string{"Hello ", "World", "?"} {
		p.AddRun().SetText(text)
	}
	if err := checkParagraphRuns(p); err == nil || !strings.Contains(err.Error(), "text getter") {
		t.Errorf("wrong concatenation passed: %v", err)
	}
	_ = doc.Close()
	for _, tc := range []struct {
		name  string
		check func() error
	}{
		{"alignment name", func() error {
			return guardParagraphAlignmentRow("A justify alignment reads back justify in memory", []string{"a new Word paragraph", "its alignment is set to justify", "its alignment getter equals justify"}, map[string]bool{})
		}},
		{"alignment step", func() error {
			return guardParagraphAlignmentRow("A justify alignment reads back both in memory", []string{"a new Word paragraph", "its alignment is set to both", "its alignment getter equals justify"}, map[string]bool{})
		}},
		{"alignment duplicate", func() error {
			return guardParagraphAlignmentRow("A justify alignment reads back both in memory", []string{"a new Word paragraph", "its alignment is set to both", "its alignment getter equals both"}, map[string]bool{"justify": true})
		}},
		{"spacing pair", func() error {
			return guardParagraphSpacingRow("A asymmetric paragraph reads back spacing in twips", []string{"a new Word paragraph", "spacing before is set to 240 and after to 240", "its before and after getters equal 240 and 120"}, map[string]bool{})
		}},
		{"spacing duplicate", func() error {
			return guardParagraphSpacingRow("A zero paragraph reads back spacing in twips", []string{"a new Word paragraph", "spacing before is set to 0 and after to 0", "its before and after getters equal 0 and 0"}, map[string]bool{"zero": true})
		}},
		{"toggle step", func() error {
			return guardParagraphSingleCase(paragraphTogglesCaseID, "Three paragraph flags read true after being enabled", []string{"a new Word paragraph", "KeepLines, PageBreakBefore and WidowControl are set true", "two getters are true in memory"})
		}},
		{"runs extra step", func() error {
			return guardParagraphSingleCase(paragraphRunsCaseID, "Three added runs concatenate in paragraph text", []string{"a new Word paragraph", "runs containing Hello-space, World and exclamation are appended in order", "the paragraph has three runs and its text equals Hello World!", "the document is saved and reopened"})
		}},
	} {
		if err := tc.check(); err == nil {
			t.Errorf("%s passed exact guard", tc.name)
		}
	}
	if err := checkParagraphAlignment(nil, "left"); err == nil {
		t.Error("missing paragraph passed alignment")
	}
	if err := checkParagraphSpacing(nil, 0, 0); err == nil {
		t.Error("missing paragraph passed spacing")
	}
	if err := checkParagraphToggles(nil); err == nil {
		t.Error("missing paragraph passed toggles")
	}
	if err := checkParagraphRuns(nil); err == nil {
		t.Error("missing paragraph passed runs")
	}
}
