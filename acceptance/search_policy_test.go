package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func searchPolicySteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var matches []*document.TextTarget
	var argumentError error
	setup := func(texts ...string) error {
		var body bytes.Buffer
		for _, text := range texts {
			body.WriteString(`<w:p><w:r><w:t xml:space="preserve">`)
			_ = xml.EscapeText(&body, []byte(text))
			body.WriteString(`</w:t></w:r></w:p>`)
		}
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body.String()+`</w:body></w:document>`))
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
		matches = nil
		argumentError = nil
		return ctx, nil
	})
	sc.Step(`^a Word paragraph containing "([^"]+)"$`, func(text string) error { return setup(text) })
	sc.Step(`^Word search paragraphs "([^"]+)" and "([^"]+)"$`, func(a, b string) error { return setup(a, b) })
	sc.Step(`^I search with normalisation for "([^"]+)"$`, func(needle string) error {
		var err error
		matches, err = s.Search(needle, document.SearchOptions{Normalized: true})
		return err
	})
	sc.Step(`^the search evidence is exactly "([^"]+)"$`, func(text string) error {
		if len(matches) != 1 || matches[0].Text() != text {
			return fmt.Errorf("expected original evidence %q", text)
		}
		return nil
	})
	sc.Step(`^that selected evidence can authorise an ordinary replacement$`, func() error { return s.Replace(matches[0], "New cost") })
	sc.Step(`^normalised search returns no partial-character target$`, func() error {
		if len(matches) != 0 {
			return fmt.Errorf("partial casefold character matched")
		}
		return nil
	})
	sc.Step(`^I rank matches for "([^"]+)" near "([^"]+)"$`, func(text, near string) error {
		var err error
		matches, err = s.Search(text, document.SearchOptions{Near: near})
		return err
	})
	sc.Step(`^both candidates remain and the contextual paragraph ranks first$`, func() error {
		if len(matches) != 2 {
			return fmt.Errorf("lost candidates")
		}
		if err := s.Replace(matches[0], "blue"); err != nil {
			return err
		}
		o, err := s.Outline(document.CurrentView)
		if err != nil {
			return err
		}
		if o.Blocks[0].Text != "red far" || o.Blocks[1].Text != "blue context" {
			return fmt.Errorf("wrong ranked identity %+v", o.Blocks)
		}
		return nil
	})
	sc.Step(`^I combine an explicit occurrence with contextual ranking$`, func() error {
		_, argumentError = s.Search("red", document.SearchOptions{Near: "red", Nth: 1})
		return nil
	})
	sc.Step(`^the search argument conflict is reported$`, func() error {
		if argumentError == nil {
			return fmt.Errorf("conflicting options accepted")
		}
		return nil
	})
}
