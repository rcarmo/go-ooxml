package acceptance

import (
	"context"
	"fmt"
	"path/filepath"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

// Only the three authored runs and named direct properties are checked after
// reopening. This does not assert Office rendering or effective formatting.
func runFormattingReadbackSteps(sc *godog.ScenarioContext, outputDir string, tableReadback *tableTextReadbackState) {
	var doc document.Document
	var runs []document.Run
	var savedPath string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, runs, savedPath = nil, nil, ""
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a new Word paragraph with three runs Bold-space, Italic-space and Colored$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		p := doc.AddParagraph()
		for _, text := range []string{"Bold ", "Italic ", "Colored"} {
			r := p.AddRun()
			r.SetText(text)
			runs = append(runs, r)
		}
		return nil
	})
	sc.Step(`^the first run is bold, the second italic, and the third has colour FF0000, font size 14 and font Arial$`, func() error {
		if len(runs) != 3 {
			return fmt.Errorf("expected three authored runs, got %d", len(runs))
		}
		runs[0].SetBold(true)
		runs[1].SetItalic(true)
		runs[2].SetColor("FF0000")
		runs[2].SetFontSize(14)
		runs[2].SetFontName("Arial")
		return nil
	})
	sc.Step(`^the document is saved and reopened$`, func() error {
		if tableReadback.doc != nil {
			return tableReadback.saveAndReopen(outputDir)
		}
		if doc == nil {
			return fmt.Errorf("no Word document")
		}
		savedPath = filepath.Join(outputDir, "selected-formatting.docx")
		if err := doc.SaveAs(savedPath); err != nil {
			return err
		}
		if err := doc.Close(); err != nil {
			return err
		}
		doc = nil
		var err error
		doc, err = document.Open(savedPath)
		return err
	})
	sc.Step(`^at least one paragraph and three runs are readable$`, func() error {
		var err error
		runs, err = reopenedFormattingRuns(doc)
		return err
	})
	sc.Step(`^the first run is bold and the second italic$`, func() error {
		return checkReopenedRunFlags(runs)
	})
	sc.Step(`^the third run reports colour FF0000, font size 14 and font Arial$`, func() error {
		return checkReopenedThirdRun(runs)
	})
}

func reopenedFormattingRuns(doc document.Document) ([]document.Run, error) {
	if doc == nil || len(doc.Paragraphs()) < 1 {
		return nil, fmt.Errorf("no paragraph after reopening")
	}
	runs := doc.Paragraphs()[0].Runs()
	if len(runs) < 3 {
		return nil, fmt.Errorf("expected at least three runs, got %d", len(runs))
	}
	return runs, nil
}

func checkReopenedRunFlags(runs []document.Run) error {
	if len(runs) < 3 || !runs[0].Bold() || !runs[1].Italic() {
		return fmt.Errorf("reopened first Bold or second Italic getter differs")
	}
	return nil
}

func checkReopenedThirdRun(runs []document.Run) error {
	if len(runs) < 3 || runs[2].Color() != "FF0000" || runs[2].FontSize() != 14 || runs[2].FontName() != "Arial" {
		return fmt.Errorf("reopened third-run colour/size/font differs")
	}
	return nil
}

func TestSelectedFormattingReadbackNegativeControls(t *testing.T) {
	for _, tc := range []struct {
		name   string
		index  int
		change func(document.Run)
		check  func([]document.Run) error
	}{
		{"first-bold", 0, func(r document.Run) { r.SetBold(false) }, checkReopenedRunFlags},
		{"second-italic", 1, func(r document.Run) { r.SetItalic(false) }, checkReopenedRunFlags},
		{"third-colour", 2, func(r document.Run) { r.SetColor("00FF00") }, checkReopenedThirdRun},
		{"third-size", 2, func(r document.Run) { r.SetFontSize(12) }, checkReopenedThirdRun},
		{"third-font", 2, func(r document.Run) { r.SetFontName("Verdana") }, checkReopenedThirdRun},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			p := doc.AddParagraph()
			r1 := p.AddRun()
			r1.SetText("Bold ")
			r1.SetBold(true)
			r2 := p.AddRun()
			r2.SetText("Italic ")
			r2.SetItalic(true)
			r3 := p.AddRun()
			r3.SetText("Colored")
			r3.SetColor("FF0000")
			r3.SetFontSize(14)
			r3.SetFontName("Arial")
			path := filepath.Join(t.TempDir(), "control.docx")
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
			runs, err := reopenedFormattingRuns(reopened)
			if err != nil {
				t.Fatal(err)
			}
			if err := checkReopenedRunFlags(runs); err != nil {
				t.Fatalf("positive flags: %v", err)
			}
			if err := checkReopenedThirdRun(runs); err != nil {
				t.Fatalf("positive third run: %v", err)
			}
			tc.change(runs[tc.index])
			changedPath := filepath.Join(t.TempDir(), "changed.docx")
			if err := reopened.SaveAs(changedPath); err != nil {
				t.Fatal(err)
			}
			if err := reopened.Close(); err != nil {
				t.Fatal(err)
			}
			changed, err := document.Open(changedPath)
			if err != nil {
				t.Fatal(err)
			}
			defer changed.Close()
			changedRuns, err := reopenedFormattingRuns(changed)
			if err != nil {
				t.Fatal(err)
			}
			if err := tc.check(changedRuns); err == nil {
				t.Fatalf("%s changed saved value did not fail its getter", tc.name)
			}
		})
	}
}
