package acceptance

import (
	"context"
	"fmt"
	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/formula"
)

func remapSteps(sc *godog.ScenarioContext) {
	var source, owner, output string
	var failure error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = ""
		owner = ""
		output = ""
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a formula "([^"]+)" located on "([^"]+)"$`, func(text, sheet string) error { source, owner = text, sheet; return nil })
	sc.Step(`^I insert (\d+) "([^"]+)" before coordinate (\d+) on "([^"]+)"$`, func(count int, axis string, at int, sheet string) error {
		var err error
		output, err = formula.InsertReferences(source, owner, formula.Insertion{Sheet: sheet, Axis: axis, At: at, Count: count})
		return err
	})
	sc.Step(`^the remapped formula is "([^"]+)"$`, func(want string) error {
		if output != want {
			return fmt.Errorf("got %q want %q", output, want)
		}
		return nil
	})
	sc.Step(`^I attempt a row insertion that shifts it beyond the grid$`, func() error {
		output, failure = formula.InsertReferences(source, owner, formula.Insertion{Sheet: "Main", Axis: "row", At: 1, Count: 1})
		return nil
	})
	sc.Step(`^remapping refuses and returns no partial expression$`, func() error {
		if failure == nil || output != "" {
			return fmt.Errorf("partial/successful remap %q %v", output, failure)
		}
		return nil
	})
}
