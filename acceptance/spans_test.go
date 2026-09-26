package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func spanSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var spans []*document.TextTarget
	var needle string
	setup := func(first, second string, paragraphs bool) error {
		esc := func(text string) string { var b bytes.Buffer; _ = xml.EscapeText(&b, []byte(text)); return b.String() }
		run := func(text, property string) string {
			return `<w:r><w:rPr><w:` + property + `/></w:rPr><w:t xml:space="preserve">` + esc(text) + `</w:t></w:r>`
		}
		body := `<w:p>` + run(first, "b") + run(second, "i") + `</w:p>`
		if paragraphs {
			body = `<w:p>` + run(first, "b") + `</w:p><w:p>` + run(second, "i") + `</w:p>`
		}
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body+`</w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		var err error
		s, err = document.OpenEditing(b.Bytes(), packaging.Limits{})
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		spans = nil
		needle = ""
		return ctx, nil
	})
	sc.Step(`^a Word paragraph with run text "([^"]+)" and "([^"]+)"$`, func(a, b string) error { return setup(a, b, false) })
	sc.Step(`^Word paragraphs containing "([^"]+)" and "([^"]+)"$`, func(a, b string) error { return setup(a, b, true) })
	sc.Step(`^I find exact Word spans for "([^"]+)"$`, func(text string) error { needle = text; var err error; spans, err = s.FindText(text); return err })
	sc.Step(`^exactly one span reports the original text "([^"]+)"$`, func(text string) error {
		if len(spans) != 1 || spans[0].Text() != text {
			return fmt.Errorf("expected unique %q, got %d spans", text, len(spans))
		}
		target, err := s.FindOne(text)
		if err != nil {
			return err
		}
		if target.Text() != text {
			return fmt.Errorf("single target text mismatch")
		}
		return nil
	})
	sc.Step(`^all three spans remain available in source order$`, func() error {
		if len(spans) != 3 {
			return fmt.Errorf("expected three spans, got %d", len(spans))
		}
		for _, span := range spans {
			if span.Text() != needle {
				return fmt.Errorf("changed evidence")
			}
		}
		return nil
	})
	sc.Step(`^single-target lookup reports ambiguous-target$`, func() error {
		_, err := s.FindOne(needle)
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "ambiguous_target" {
			return fmt.Errorf("expected ambiguity, got %v", err)
		}
		return nil
	})
	sc.Step(`^no spans are returned$`, func() error {
		if len(spans) != 0 {
			return fmt.Errorf("unexpected spans")
		}
		return nil
	})
}
