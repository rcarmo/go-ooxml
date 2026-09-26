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

func runBoundarySteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var source []byte
	var target *document.TextTarget
	var failure error
	setup := func(condition string) error {
		property, attribute, marker := `<w:b/>`, "", ""
		if condition == "different direct formatting" {
			property = `<w:i/>`
		}
		if condition == "different run attributes" {
			attribute = ` w:rsidR="1234"`
		}
		if condition == "an intervening bookmark" {
			marker = `<w:bookmarkStart w:id="1" w:name="mark"/>`
		}
		body := `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t></w:r>` + marker + `<w:r` + attribute + `><w:rPr>` + property + `</w:rPr><w:t>cd</w:t></w:r></w:p>`
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body+`</w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = document.OpenEditing(source, packaging.Limits{})
		if err != nil {
			return err
		}
		target, err = s.FindOne("abcd")
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		target = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^adjacent Word runs with identical direct formatting$`, func() error { return setup("") })
	sc.Step(`^adjacent Word runs with "([^"]+)"$`, setup)
	sc.Step(`^I insert a character at their shared text boundary$`, func() error { return s.Replace(target, "abXcd") })
	sc.Step(`^the combined text includes the insertion with unchanged run properties$`, func() error {
		o, err := s.Outline(document.CurrentView)
		if err != nil {
			return err
		}
		if len(o.Blocks) != 1 || o.Blocks[0].Text != "abXcd" {
			return fmt.Errorf("wrong text %+v", o.Blocks)
		}
		return nil
	})
	sc.Step(`^I attempt an insertion at their shared text boundary$`, func() error { failure = s.Replace(target, "abXcd"); return nil })
	sc.Step(`^the insertion refuses and retains the original archive$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected refusal, got %v", failure)
		}
		dir, err := os.MkdirTemp("", "boundary-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.docx")
		if _, err = s.SaveAs(path); err != nil {
			return err
		}
		b, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		if !bytes.Equal(source, b) {
			return fmt.Errorf("refusal mutated archive")
		}
		return nil
	})
}
