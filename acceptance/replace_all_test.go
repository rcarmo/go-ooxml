package acceptance

import (
	"bytes"
	"context"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func replaceAllSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var original []byte
	var result document.ReplaceAllResult
	setup := func() error {
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body><w:p><w:r><w:t>red red</w:t></w:r></w:p><w:p><w:hyperlink><w:r><w:t>red</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		original = b.Bytes()
		var err error
		s, err = document.OpenEditing(original, packaging.Limits{})
		return err
	}
	unchanged := func() error {
		dir, err := os.MkdirTemp("", "replace-all-")
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
		if !bytes.Equal(original, b) {
			return fmt.Errorf("unexpected package mutation")
		}
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		result = document.ReplaceAllResult{}
		return ctx, nil
	})
	sc.Step(`^two plain Word matches and one hyperlink-owned match$`, setup)
	sc.Step(`^I replace all occurrences of "([^"]+)" with "([^"]+)"$`, func(needle, replacement string) error {
		var err error
		result, err = s.ReplaceAll(needle, replacement, false)
		return err
	})
	sc.Step(`^two matches change and one unsupported match is reported$`, func() error {
		if result.Matched != 3 || result.Changed != 2 || len(result.Refusals) != 1 || result.Refusals[0].Kind != "unsupported_structure" {
			return fmt.Errorf("report %+v", result)
		}
		return nil
	})
	sc.Step(`^the hyperlink-owned text remains unchanged$`, func() error {
		o, err := s.Outline(document.CurrentView)
		if err != nil {
			return err
		}
		if len(o.Blocks) != 2 || o.Blocks[0].Text != "blue blue" || o.Blocks[1].Text != "red" {
			return fmt.Errorf("outline %+v", o.Blocks)
		}
		return nil
	})
	sc.Step(`^all three matches are skipped without refusals or mutation$`, func() error {
		if result.Matched != 3 || result.Skipped != 3 || result.Changed != 0 || len(result.Refusals) != 0 {
			return fmt.Errorf("report %+v", result)
		}
		return unchanged()
	})
	sc.Step(`^I replace all occurrences with an XML-illegal value$`, func() error { var err error; result, err = s.ReplaceAll("red", "\x00", false); return err })
	sc.Step(`^no match changes and the original package bytes remain exact$`, func() error {
		if result.Changed != 0 || len(result.Refusals) != 3 {
			return fmt.Errorf("report %+v", result)
		}
		return unchanged()
	})
}
