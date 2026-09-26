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

func searchScopeSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var matches []*document.TextTarget
	var source []byte
	setup := func(body, header string) error {
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body+`</w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		if header != "" {
			_, _ = q.AddPart("word/header.xml", packaging.ContentTypeHeader, []byte(`<w:hdr xmlns:w="`+packaging.NSWordprocessingML+`"><w:p><w:r><w:t>`+header+`</w:t></w:r></w:p></w:hdr>`))
			q.AddRelationship("word/document.xml", "header.xml", packaging.RelTypeHeader)
		}
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = document.OpenEditing(source, packaging.Limits{})
		return err
	}
	refuse := func() error {
		if len(matches) != 1 {
			return fmt.Errorf("match absent")
		}
		err := s.Replace(matches[0], "new")
		var r *packaging.Refusal
		if !errors.As(err, &r) {
			return fmt.Errorf("expected safe refusal, got %v", err)
		}
		dir, err := os.MkdirTemp("", "scope-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		p := filepath.Join(dir, "out.docx")
		if _, err = s.SaveAs(p); err != nil {
			return err
		}
		b, err := os.ReadFile(p)
		if err != nil {
			return err
		}
		if !bytes.Equal(source, b) {
			return fmt.Errorf("refusal changed package")
		}
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		matches = nil
		return ctx, nil
	})
	sc.Step(`^two scoped Word paragraphs containing "first" and "second"$`, func() error {
		return setup(`<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>`, "")
	})
	sc.Step(`^I search for their text with an explicit paragraph separator$`, func() error {
		var err error
		matches, err = s.Search("first\nsecond", document.SearchOptions{})
		return err
	})
	sc.Step(`^one cross-paragraph match reports the exact combined text$`, func() error {
		if len(matches) != 1 || matches[0].Text() != "first\nsecond" {
			return fmt.Errorf("cross-paragraph match missing")
		}
		return nil
	})
	sc.Step(`^attempting to edit that cross-paragraph match refuses$`, refuse)
	sc.Step(`^scoped Word text with a tracked deletion$`, func() error {
		return setup(`<w:p><w:del><w:r><w:delText>removed</w:delText></w:r></w:del><w:ins><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>`, "")
	})
	sc.Step(`^I search original-view text for "removed"$`, func() error {
		var err error
		matches, err = s.Search("removed", document.SearchOptions{View: document.OriginalView})
		return err
	})
	sc.Step(`^the historical match reports "removed"$`, func() error {
		if len(matches) != 1 || matches[0].Text() != "removed" {
			return fmt.Errorf("historical evidence missing")
		}
		return nil
	})
	sc.Step(`^editing the historical match refuses without mutation$`, refuse)
	sc.Step(`^scoped Word body and header text both equal to "repeat"$`, func() error { return setup(`<w:p><w:r><w:t>repeat</w:t></w:r></w:p>`, "repeat") })
	sc.Step(`^I search only the header story for "repeat"$`, func() error {
		var err error
		matches, err = s.Search("repeat", document.SearchOptions{Story: "word/header.xml"})
		return err
	})
	sc.Step(`^exactly one match identifies the header story$`, func() error {
		if len(matches) != 1 || matches[0].Story() != "word/header.xml" {
			return fmt.Errorf("story identity missing")
		}
		return nil
	})
}
