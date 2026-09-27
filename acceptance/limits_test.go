package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func limitSteps(sc *godog.ScenarioContext) {
	var source []byte
	var limits packaging.Limits
	var got *packaging.Package
	var result error
	setup := func() error {
		p := packaging.New()
		if _, err := p.AddPart("data.xml", packaging.ContentTypeXML, []byte(`<data/>`)); err != nil {
			return err
		}
		var b bytes.Buffer
		if err := p.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		limits = packaging.Limits{MaxSourceBytes: 1 << 20, MaxEntries: 10, MaxPartBytes: 1 << 20, MaxTotalBytes: 1 << 20}
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		got = nil
		result = nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if got != nil {
			_ = got.Close()
		}
		return ctx, nil
	})
	sc.Step(`^an archive exceeding its "([^"]+)" budget$`, func(budget string) error {
		if err := setup(); err != nil {
			return err
		}
		switch budget {
		case "source":
			limits.MaxSourceBytes = 1
		case "entries":
			limits.MaxEntries = 1
		case "part":
			limits.MaxPartBytes = 1
		case "total":
			limits.MaxTotalBytes = 1
		default:
			return fmt.Errorf("unknown budget")
		}
		return nil
	})
	sc.Step(`^an archive within all configured budgets$`, setup)
	sc.Step(`^I open it with the configured resource limits$`, func() error {
		got, result = packaging.OpenReaderWithLimits(bytes.NewReader(source), int64(len(source)), limits)
		return nil
	})
	sc.Step(`^I receive a typed resource-limit refusal$`, func() error {
		var refusal *packaging.Refusal
		if !errors.As(result, &refusal) || refusal.Kind != "resource_limit" {
			return fmt.Errorf("expected resource_limit, got %v", result)
		}
		if got != nil {
			return fmt.Errorf("partial package returned")
		}
		return nil
	})
	sc.Step(`^its XML payload is available without changes$`, func() error {
		if result != nil {
			return result
		}
		p, err := got.GetPart("data.xml")
		if err != nil {
			return err
		}
		data, err := p.Content()
		if err != nil {
			return err
		}
		if string(data) != `<data/>` {
			return fmt.Errorf("changed payload")
		}
		return nil
	})
}
