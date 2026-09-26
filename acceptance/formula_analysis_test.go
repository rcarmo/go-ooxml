package acceptance

import (
	"context"
	"fmt"
	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/formula"
)

func formulaSteps(sc *godog.ScenarioContext) {
	var source string
	var references []formula.Reference
	var failure error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = ""
		references = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^the formula expression "([^"]+)"$`, func(text string) error { source = text; return nil })
	sc.Step(`^I analyse its static references$`, func() error { references, failure = formula.Analyze(source); return nil })
	sc.Step(`^exactly (\d+) reference ranges are reported$`, func(count int) error {
		if failure != nil {
			return failure
		}
		if len(references) != count {
			return fmt.Errorf("got%d want%d", len(references), count)
		}
		return nil
	})
	sc.Step(`^dependency analysis refuses instead of returning a partial answer$`, func() error {
		if failure == nil || len(references) != 0 {
			return fmt.Errorf("partial or successful unknown analysis")
		}
		return nil
	})
}
