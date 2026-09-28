package acceptance

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

// This is an in-memory Go API predicate, not a claim that the simultaneous
// effects form a valid saved WordprocessingML run.
func runEffectsSteps(sc *godog.ScenarioContext) {
	var doc document.Document
	var run document.Run
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, run = nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a new Word run$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		run = doc.AddParagraph().AddRun()
		return nil
	})
	sc.Step(`^DoubleStrike, Caps, SmallCaps, Outline, Shadow, Emboss, Imprint and Vanish are set true$`, func() error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		setAllRunEffects(run)
		return nil
	})
	sc.Step(`^all eight corresponding getters return true in memory$`, func() error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		return checkAllRunEffects(run)
	})
	sc.Step(`^its underline style is set to (single|double|thick|dotted|dash|wave)$`, func(style string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		run.SetUnderlineStyle(style)
		return nil
	})
	sc.Step(`^Underline is true and UnderlineStyle equals (single|double|thick|dotted|dash|wave)$`, func(style string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		return checkRunUnderlineStyle(run, style)
	})
	sc.Step(`^its font name is set to (Arial|Times New Roman|Calibri|Courier New|Georgia|Verdana)$`, func(font string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		run.SetFontName(font)
		return nil
	})
	sc.Step(`^its font-name getter equals (Arial|Times New Roman|Calibri|Courier New|Georgia|Verdana)$`, func(font string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		return checkRunFontName(run, font)
	})
	sc.Step(`^its colour is set to (FF0000|#FF0000|ff0000)$`, func(colour string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		run.SetColor(colour)
		return nil
	})
	sc.Step(`^its in-memory colour getter equals (FF0000|ff0000)$`, func(want string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		return checkRunColor(run, want)
	})
	sc.Step(`^highlight is set to (yellow|cyan|darkBlue|lightGray|black)$`, func(colour string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		run.SetHighlight(colour)
		return nil
	})
	sc.Step(`^the Highlight getter equals (yellow|cyan|darkBlue|lightGray|black)$`, func(want string) error {
		if run == nil {
			return fmt.Errorf("no Word run")
		}
		return checkRunHighlight(run, want)
	})
}

func checkRunColor(run document.Run, want string) error {
	if got := run.Color(); got != want {
		return fmt.Errorf("Color getter = %q, want %q", got, want)
	}
	return nil
}

func checkRunHighlight(run document.Run, want string) error {
	if got := run.Highlight(); got != want {
		return fmt.Errorf("Highlight getter = %q, want %q", got, want)
	}
	return nil
}

func checkRunUnderlineStyle(run document.Run, want string) error {
	if !run.Underline() {
		return fmt.Errorf("Underline getter is false for %s", want)
	}
	if got := run.UnderlineStyle(); got != want {
		return fmt.Errorf("UnderlineStyle getter = %q, want %q", got, want)
	}
	return nil
}

func checkRunFontName(run document.Run, want string) error {
	if got := run.FontName(); got != want {
		return fmt.Errorf("FontName getter = %q, want %q", got, want)
	}
	return nil
}

func setAllRunEffects(run document.Run) {
	run.SetDoubleStrike(true)
	run.SetCaps(true)
	run.SetSmallCaps(true)
	run.SetOutline(true)
	run.SetShadow(true)
	run.SetEmboss(true)
	run.SetImprint(true)
	run.SetVanish(true)
}

func checkAllRunEffects(run document.Run) error {
	for _, effect := range []struct {
		name string
		read func() bool
	}{
		{"DoubleStrike", run.DoubleStrike},
		{"Caps", run.Caps},
		{"SmallCaps", run.SmallCaps},
		{"Outline", run.Outline},
		{"Shadow", run.Shadow},
		{"Emboss", run.Emboss},
		{"Imprint", run.Imprint},
		{"Vanish", run.Vanish},
	} {
		if !effect.read() {
			return fmt.Errorf("%s getter is not true in memory", effect.name)
		}
	}
	return nil
}

func TestRunFormattingGetterNegativeControls(t *testing.T) {
	for _, style := range []string{"single", "double", "thick", "dotted", "dash", "wave"} {
		t.Run("underline/"+style, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			run := doc.AddParagraph().AddRun()
			run.SetUnderlineStyle(style)
			if err := checkRunUnderlineStyle(run, style); err != nil {
				t.Fatalf("positive control: %v", err)
			}
			run.SetUnderlineStyle("none")
			if err := checkRunUnderlineStyle(run, style); err == nil || !strings.Contains(err.Error(), "Underline") {
				t.Fatalf("cleared Underline should fail %s: %v", style, err)
			}
			run.SetUnderlineStyle("single")
			if style == "single" {
				run.SetUnderlineStyle("double")
			}
			if err := checkRunUnderlineStyle(run, style); err == nil || !strings.Contains(err.Error(), "UnderlineStyle") {
				t.Fatalf("wrong UnderlineStyle should fail %s: %v", style, err)
			}
		})
	}
	for _, font := range []string{"Arial", "Times New Roman", "Calibri", "Courier New", "Georgia", "Verdana"} {
		t.Run("font/"+font, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			run := doc.AddParagraph().AddRun()
			run.SetFontName(font)
			if err := checkRunFontName(run, font); err != nil {
				t.Fatalf("positive control: %v", err)
			}
			wrong := "Arial"
			if font == wrong {
				wrong = "Verdana"
			}
			run.SetFontName(wrong)
			if err := checkRunFontName(run, font); err == nil || !strings.Contains(err.Error(), "FontName") {
				t.Fatalf("wrong FontName should fail %s: %v", font, err)
			}
		})
	}
}

func TestRunColourHighlightNegativeControls(t *testing.T) {
	for _, tc := range []struct{ input, want string }{{"FF0000", "FF0000"}, {"#FF0000", "FF0000"}, {"ff0000", "ff0000"}} {
		t.Run("colour/"+tc.input, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			run := doc.AddParagraph().AddRun()
			run.SetColor(tc.input)
			if err := checkRunColor(run, tc.want); err != nil {
				t.Fatalf("positive control: %v", err)
			}
			run.SetColor("00FF00")
			if err := checkRunColor(run, tc.want); err == nil || !strings.Contains(err.Error(), "Color") {
				t.Fatalf("wrong colour should fail %s: %v", tc.input, err)
			}
		})
	}
	for _, colour := range []string{"yellow", "cyan", "darkBlue", "lightGray", "black"} {
		t.Run("highlight/"+colour, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			run := doc.AddParagraph().AddRun()
			run.SetHighlight(colour)
			if err := checkRunHighlight(run, colour); err != nil {
				t.Fatalf("positive control: %v", err)
			}
			run.SetHighlight("")
			if err := checkRunHighlight(run, colour); err == nil || !strings.Contains(err.Error(), "Highlight") {
				t.Fatalf("cleared highlight should fail %s: %v", colour, err)
			}
		})
	}
}

func TestRunEffectsGetterNegativeControls(t *testing.T) {
	for _, effect := range []struct {
		name  string
		clear func(document.Run)
	}{
		{"DoubleStrike", func(r document.Run) { r.SetDoubleStrike(false) }},
		{"Caps", func(r document.Run) { r.SetCaps(false) }},
		{"SmallCaps", func(r document.Run) { r.SetSmallCaps(false) }},
		{"Outline", func(r document.Run) { r.SetOutline(false) }},
		{"Shadow", func(r document.Run) { r.SetShadow(false) }},
		{"Emboss", func(r document.Run) { r.SetEmboss(false) }},
		{"Imprint", func(r document.Run) { r.SetImprint(false) }},
		{"Vanish", func(r document.Run) { r.SetVanish(false) }},
	} {
		t.Run(effect.name, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			run := doc.AddParagraph().AddRun()
			setAllRunEffects(run)
			if err := checkAllRunEffects(run); err != nil {
				t.Fatalf("positive control: %v", err)
			}
			effect.clear(run)
			if err := checkAllRunEffects(run); err == nil || !strings.Contains(err.Error(), effect.name) {
				t.Fatalf("clearing %s should fail its exact getter: %v", effect.name, err)
			}
		})
	}
}
