package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"os"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const negativeBudgetFixtureID = "fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5"

// A ReaderAt that records even attempted reads proves invalid configuration is
// rejected before ZIP metadata/member inspection. It does not fabricate a
// malformed archive: the sealed valid bytes are opened separately first.
type observedArchiveReader struct {
	data  []byte
	reads int
}

func (r *observedArchiveReader) ReadAt(p []byte, off int64) (int, error) {
	r.reads++
	return bytes.NewReader(r.data).ReadAt(p, off)
}

func negativeBudgetSteps(sc *godog.ScenarioContext) {
	var source, snapshot []byte
	var observed *observedArchiveReader
	var budget string
	var got *packaging.Package
	var refusal error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, snapshot, observed, budget, got, refusal = nil, nil, nil, "", nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if got != nil {
			_ = got.Close()
		}
		return ctx, nil
	})
	sc.Step(`^the byte-sealed valid DOCX archive fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5 and a separate caller byte snapshot$`, func() error {
		path, err := testutil.LookupFixture(negativeBudgetFixtureID)
		if err != nil {
			return err
		}
		source, err = os.ReadFile(path)
		if err != nil {
			return err
		}
		if len(source) != 36563 {
			return fmt.Errorf("wrong sealed DOCX size: %d", len(source))
		}
		snapshot = bytes.Clone(source)
		// A successful independent baseline proves the negative result is due to
		// caller configuration rather than unreadable source bytes.
		valid, err := packaging.OpenReaderWithLimits(bytes.NewReader(source), int64(len(source)), packaging.Limits{})
		if err != nil || valid == nil {
			return fmt.Errorf("valid DOCX refused: %v", err)
		}
		defer valid.Close()
		if len(valid.Parts()) == 0 {
			return fmt.Errorf("valid DOCX has no parts")
		}
		return nil
	})
	sc.Step(`^only the (source bytes|entry count) admission budget is set to -1$`, func(field string) error {
		budget = field
		return nil
	})
	sc.Step(`^bounded package admission checks that archive$`, func() error {
		if !bytes.Equal(source, snapshot) || budget == "" {
			return fmt.Errorf("missing sealed input or budget")
		}
		limits := packaging.Limits{}
		switch budget {
		case "source bytes":
			limits.MaxSourceBytes = -1
		case "entry count":
			limits.MaxEntries = -1
		default:
			return fmt.Errorf("unexpected budget %q", budget)
		}
		observed = &observedArchiveReader{data: source}
		got, refusal = packaging.OpenReaderWithLimits(observed, int64(len(source)), limits)
		return nil
	})
	sc.Step(`^it refuses the invalid caller budget before reading source metadata or ZIP members and returns no package or parts$`, func() error {
		if observed == nil || observed.reads != 0 || got != nil || refusal == nil {
			return fmt.Errorf("read count=%d package=%v error=%v", func() int {
				if observed == nil {
					return -1
				}
				return observed.reads
			}(), got, refusal)
		}
		return nil
	})
	sc.Step(`^the refusal is an invalid-argument result, not a resource-limit or malformed-archive result$`, func() error {
		var typed *packaging.Refusal
		if refusal == nil || refusal.Error() != "negative archive size or resource budget" || errors.As(refusal, &typed) {
			return fmt.Errorf("wrong invalid-argument result: %v", refusal)
		}
		return nil
	})
	sc.Step(`^the caller's archive bytes remain unchanged$`, func() error {
		if !bytes.Equal(source, snapshot) || got != nil {
			return fmt.Errorf("caller bytes changed or package delivered")
		}
		return nil
	})
}
