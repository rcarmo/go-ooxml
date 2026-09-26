package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func spanBatchSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var spans []*document.TextTarget
	var original, before []byte
	var result error
	setup := func(wrapper bool) error {
		body := `<w:p><w:r><w:t xml:space="preserve">red red red</w:t></w:r></w:p>`
		if wrapper {
			body = `<w:p><w:r><w:t>red</w:t></w:r></w:p><w:p><w:hyperlink><w:r><w:t>red</w:t></w:r></w:hyperlink></w:p>`
		}
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body+`</w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		original = b.Bytes()
		var err error
		s, err = document.OpenEditing(original, packaging.Limits{})
		if err != nil {
			return err
		}
		spans, err = s.FindText("red")
		return err
	}
	save := func() ([]byte, error) {
		dir, err := os.MkdirTemp("", "word-batch-")
		if err != nil {
			return nil, err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.docx")
		if _, err = s.SaveAs(path); err != nil {
			return nil, err
		}
		return os.ReadFile(path)
	}
	apply := func(replacement string) error {
		changes := make([]document.TextReplacement, len(spans))
		for i, span := range spans {
			changes[i] = document.TextReplacement{Target: span, Text: replacement}
		}
		result = s.ReplaceBatch(changes)
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		spans = nil
		result = nil
		before = nil
		return ctx, nil
	})
	sc.Step(`^selected Word spans for three occurrences of "red"$`, func() error { return setup(false) })
	sc.Step(`^selected Word spans with a later hyperlink-owned occurrence$`, func() error { return setup(true) })
	sc.Step(`^I replace all selected spans with "([^"]+)"$`, func(text string) error { _ = apply(text); return result })
	sc.Step(`^the paragraph reads "([^"]+)"$`, func(text string) error {
		o, err := s.Outline(document.CurrentView)
		if err != nil {
			return err
		}
		if len(o.Blocks) != 1 || o.Blocks[0].Text != text {
			return fmt.Errorf("outline %+v", o.Blocks)
		}
		return nil
	})
	sc.Step(`^every changed selected target is consumed$`, func() error {
		for _, span := range spans {
			err := s.Replace(span, "again")
			var r *packaging.Refusal
			if !errors.As(err, &r) || r.Kind != "stale_target" {
				return fmt.Errorf("got %v", err)
			}
		}
		return nil
	})
	sc.Step(`^I attempt a selected batch replacement with "([^"]+)"$`, apply)
	sc.Step(`^the selected batch refuses without changing the source archive$`, func() error {
		var r *packaging.Refusal
		if !errors.As(result, &r) {
			return fmt.Errorf("expected refusal, got %v", result)
		}
		got, err := save()
		if err != nil {
			return err
		}
		if !bytes.Equal(got, original) {
			return fmt.Errorf("partial batch mutation")
		}
		return nil
	})
	sc.Step(`^the first ordinary target remains reusable$`, func() error { return s.Replace(spans[0], "green") })
	sc.Step(`^the first occurrence has already been changed$`, func() error {
		if err := s.Replace(spans[0], "green"); err != nil {
			return err
		}
		var err error
		before, err = save()
		return err
	})
	sc.Step(`^the stale batch leaves the previously committed edit unchanged$`, func() error {
		var r *packaging.Refusal
		if !errors.As(result, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("expected stale, got %v", result)
		}
		after, err := save()
		if err != nil {
			return err
		}
		if !bytes.Equal(before, after) {
			return fmt.Errorf("stale batch modified committed state")
		}
		return nil
	})
}
