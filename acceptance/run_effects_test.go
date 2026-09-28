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
