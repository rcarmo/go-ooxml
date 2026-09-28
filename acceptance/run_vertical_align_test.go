package acceptance

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
)

// The shared case checks direct Go getters on separate in-memory runs.
func runVerticalAlignSteps(sc *godog.ScenarioContext) {
	var doc document.Document
	var first, second document.Run
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, first, second = nil, nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	sc.Step(`^two new Word runs$`, func() error {
		var err error
		doc, err = document.New()
		if err != nil {
			return err
		}
		paragraph := doc.AddParagraph()
		first = paragraph.AddRun()
		second = paragraph.AddRun()
		return nil
	})
	sc.Step(`^Superscript is enabled on the first and Subscript on the second$`, func() error {
		if first == nil || second == nil {
			return fmt.Errorf("two Word runs required")
		}
		first.SetSuperscript(true)
		second.SetSubscript(true)
		return nil
	})
	sc.Step(`^the first reports superscript true and subscript false$`, func() error {
		return checkRunVerticalAlign(first, "first", true, false)
	})
	sc.Step(`^the second reports subscript true and superscript false$`, func() error {
		return checkRunVerticalAlign(second, "second", false, true)
	})
}

func checkVerticalFlags(name string, superscript, subscript, wantSuperscript, wantSubscript bool) error {
	if superscript != wantSuperscript {
		return fmt.Errorf("%s Superscript getter = %t, want %t", name, superscript, wantSuperscript)
	}
	if subscript != wantSubscript {
		return fmt.Errorf("%s Subscript getter = %t, want %t", name, subscript, wantSubscript)
	}
	return nil
}

func checkRunVerticalAlign(run document.Run, name string, wantSuperscript, wantSubscript bool) error {
	if run == nil {
		return fmt.Errorf("%s Word run missing", name)
	}
	return checkVerticalFlags(name, run.Superscript(), run.Subscript(), wantSuperscript, wantSubscript)
}

func TestRunVerticalAlignNegativeControls(t *testing.T) {
	doc, err := document.New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	paragraph := doc.AddParagraph()
	first, second := paragraph.AddRun(), paragraph.AddRun()
	first.SetSuperscript(true)
	second.SetSubscript(true)
	if err := checkRunVerticalAlign(first, "first", true, false); err != nil {
		t.Fatalf("first positive control: %v", err)
	}
	if err := checkRunVerticalAlign(second, "second", false, true); err != nil {
		t.Fatalf("second positive control: %v", err)
	}
	for _, tc := range []struct {
		name                           string
		super, sub, wantSuper, wantSub bool
	}{
		{"first superscript false", false, false, true, false},
		{"first subscript true", true, true, true, false},
		{"second subscript false", false, false, false, true},
		{"second superscript true", true, true, false, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if err := checkVerticalFlags(tc.name, tc.super, tc.sub, tc.wantSuper, tc.wantSub); err == nil {
				t.Fatal("altered getter flag passed")
			}
		})
	}
	first.SetSuperscript(false)
	if err := checkRunVerticalAlign(first, "first", true, false); err == nil || !strings.Contains(err.Error(), "Superscript") {
		t.Fatalf("cleared first superscript passed: %v", err)
	}
	first.SetSubscript(true)
	if err := checkRunVerticalAlign(first, "first", true, false); err == nil || !strings.Contains(err.Error(), "Superscript") {
		t.Fatalf("first switched to subscript passed: %v", err)
	}
	second.SetSubscript(false)
	if err := checkRunVerticalAlign(second, "second", false, true); err == nil || !strings.Contains(err.Error(), "Subscript") {
		t.Fatalf("cleared second subscript passed: %v", err)
	}
	second.SetSuperscript(true)
	if err := checkRunVerticalAlign(second, "second", false, true); err == nil || !strings.Contains(err.Error(), "Superscript") {
		t.Fatalf("second switched to superscript passed: %v", err)
	}
}
