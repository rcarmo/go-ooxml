package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func zip64Steps(sc *godog.ScenarioContext) {
	var source, original, out []byte
	var p *packaging.Preserved
	var failure error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		original = nil
		out = nil
		p = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a tiny ZIP64 package using "([^"]+)"$`, func(layout string) error {
		opt := testutil.ZIP64Fixture{}
		switch layout {
		case "stored local sizes":
		case "deflated local sizes":
			opt.Deflate = true
		case "signed descriptor":
			opt.Deflate = true
			opt.Descriptor = "signed"
		case "unsigned descriptor":
			opt.Deflate = true
			opt.Descriptor = "unsigned"
		default:
			return fmt.Errorf("unknown layout")
		}
		source = testutil.TinyZIP64(opt)
		original = bytes.Clone(source)
		return nil
	})
	sc.Step(`^I open and round-trip its retained payloads$`, func() error {
		var err error
		p, err = packaging.OpenPreserved(source, packaging.Limits{MaxSourceBytes: 1 << 20, MaxEntries: 10, MaxPartBytes: 1 << 16, MaxTotalBytes: 1 << 18})
		if err != nil {
			return err
		}
		var b bytes.Buffer
		if err = p.WriteTo(&b); err != nil {
			return err
		}
		out = b.Bytes()
		return nil
	})
	sc.Step(`^all ZIP64 source bytes and the empty member are preserved$`, func() error {
		if !bytes.Equal(source, out) || !bytes.Equal(source, original) {
			return fmt.Errorf("no-op bytes differ")
		}
		b, _, err := p.Part("empty.bin")
		if err != nil || len(b) != 0 {
			return fmt.Errorf("empty member %v", err)
		}
		b, _, err = p.Part("data.bin")
		if err != nil || string(b) != "zip64 payload" {
			return fmt.Errorf("payload mismatch %v", err)
		}
		return nil
	})
	sc.Step(`^a tiny ZIP64 package with "([^"]+)"$`, func(defect string) error {
		opt := testutil.ZIP64Fixture{Defect: defect}
		if defect == "descriptor mismatch" {
			opt.Descriptor = "signed"
		}
		source = testutil.TinyZIP64(opt)
		original = bytes.Clone(source)
		return nil
	})
	sc.Step(`^I validate the ZIP64 package for editing$`, func() error {
		_, failure = packaging.OpenPreserved(source, packaging.Limits{MaxSourceBytes: 1 << 20, MaxEntries: 10, MaxPartBytes: 1 << 16, MaxTotalBytes: 1 << 18})
		return nil
	})
	sc.Step(`^ZIP64 intake refuses with a typed error and unchanged bytes$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected typed refusal: %v", failure)
		}
		if !bytes.Equal(original, source) {
			return fmt.Errorf("input mutated")
		}
		return nil
	})
}
