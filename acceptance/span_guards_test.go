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

func spanGuardSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var original []byte
	var failure error
	setup := func(body string) error {
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body><w:p>`+body+`</w:p></w:body></w:document>`))
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
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^Word text split around a drawing-only run$`, func() error { return setup(`<w:r><w:t>Al</w:t></w:r><w:r><w:drawing/></w:r><w:r><w:t>pha</w:t></w:r>`) })
	sc.Step(`^boundary runs with equal property bytes but different prefix bindings$`, func() error {
		return setup(`<w:r xmlns:x="urn:first"><w:rPr><x:flag/></w:rPr><w:t>ab</w:t></w:r><w:r xmlns:x="urn:second"><w:rPr><x:flag/></w:rPr><w:t>cd</w:t></w:r>`)
	})
	sc.Step(`^I attempt the matched cross-run correction$`, func() error {
		t, err := s.FindOne("Alpha")
		if err != nil {
			return err
		}
		failure = s.Replace(t, "Omega")
		return nil
	})
	sc.Step(`^I attempt the matched boundary insertion$`, func() error {
		t, err := s.FindOne("abcd")
		if err != nil {
			return err
		}
		failure = s.Replace(t, "abXcd")
		return nil
	})
	sc.Step(`^the hidden-content correction refuses without changing the archive$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected guard refusal, got %v", failure)
		}
		dir, err := os.MkdirTemp("", "guard-")
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
		if !bytes.Equal(b, original) {
			return fmt.Errorf("refusal mutated source")
		}
		return nil
	})
}
