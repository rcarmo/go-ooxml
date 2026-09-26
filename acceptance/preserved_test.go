package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"fmt"
	"io"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func preservedSteps(sc *godog.ScenarioContext) {
	var source, original, output []byte
	var p *packaging.Preserved
	var result error
	setup := func() error {
		q := packaging.New()
		_, err := q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`))
		if err != nil {
			return err
		}
		_, err = q.AddPart("custom/opaque.bin", "application/octet-stream", []byte{0, 42, 128, 255})
		if err != nil {
			return err
		}
		var b bytes.Buffer
		if err = q.WriteTo(&b); err != nil {
			return err
		}
		source = bytes.Clone(b.Bytes())
		original = bytes.Clone(source)
		p, result = packaging.OpenPreserved(source, packaging.Limits{})
		return result
	}
	save := func() error {
		var b bytes.Buffer
		if err := p.WriteTo(&b); err != nil {
			return err
		}
		output = b.Bytes()
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		original = nil
		output = nil
		p = nil
		result = nil
		return ctx, nil
	})
	sc.Step(`^an Office source archive with an opaque custom part$`, setup)
	sc.Step(`^I open a retained-source editing session and save without edits$`, save)
	sc.Step(`^the output archive is byte-identical to the source$`, func() error {
		if !bytes.Equal(output, original) {
			return fmt.Errorf("no-op archive changed")
		}
		return nil
	})
	sc.Step(`^I replace the main XML part using its current fingerprint$`, func() error {
		_, hash, err := p.Part("word/document.xml")
		if err != nil {
			return err
		}
		result = p.Replace([]packaging.Replacement{{Part: "word/document.xml", ExpectedSHA256: hash, Data: []byte(`<document><changed/></document>`)}})
		if result != nil {
			return result
		}
		return save()
	})
	sc.Step(`^only that member payload changes after saving$`, func() error {
		before, err := zipPayloads(original)
		if err != nil {
			return err
		}
		after, err := zipPayloads(output)
		if err != nil {
			return err
		}
		if len(before) != len(after) {
			return fmt.Errorf("member set changed")
		}
		for name, b := range before {
			a, ok := after[name]
			if !ok {
				return fmt.Errorf("lost part %s", name)
			}
			if name == "word/document.xml" {
				if string(a) != `<document><changed/></document>` {
					return fmt.Errorf("replacement missing")
				}
			} else if !bytes.Equal(a, b) {
				return fmt.Errorf("unrelated payload changed %s", name)
			}
		}
		return nil
	})
	sc.Step(`^the retained source bytes remain unchanged$`, func() error {
		if !bytes.Equal(source, original) {
			return fmt.Errorf("caller source mutated")
		}
		return nil
	})
	sc.Step(`^I request two replacements with one stale fingerprint$`, func() error {
		_, hash, err := p.Part("word/document.xml")
		if err != nil {
			return err
		}
		result = p.Replace([]packaging.Replacement{{Part: "word/document.xml", ExpectedSHA256: hash, Data: []byte(`<changed/>`)}, {Part: "custom/opaque.bin", ExpectedSHA256: "stale", Data: []byte("wrong")}})
		return nil
	})
	sc.Step(`^the batch returns a stale-target refusal$`, func() error {
		var r *packaging.Refusal
		if !errors.As(result, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("expected stale_target, got %v", result)
		}
		return nil
	})
	sc.Step(`^the session still saves the exact original archive$`, func() error {
		if err := save(); err != nil {
			return err
		}
		if !bytes.Equal(output, original) {
			return fmt.Errorf("failed batch mutated package")
		}
		return nil
	})
}
func zipPayloads(data []byte) (map[string][]byte, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	out := map[string][]byte{}
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, err
		}
		out[f.Name] = b
	}
	return out, nil
}
