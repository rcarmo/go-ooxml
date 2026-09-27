package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func wordSteps(sc *godog.ScenarioContext) {
	var source, main []byte
	var session *document.EditSession
	var target *document.TextTarget
	var failure error
	var dir, out string
	setup := func(condition string) error {
		body := `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">old text</w:t></w:r></w:p>`
		switch condition {
		case "duplicate text":
			body += body
		case "tracked revisions":
			body = `<w:ins w:id="1">` + body + `</w:ins>`
		case "field boundary":
			body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:t>old text</w:t></w:r></w:p>`
		case "significant whitespace without preserve":
			body = `<w:p><w:r><w:t>old text</w:t></w:r></w:p>`
		}
		main = []byte(`<w:document xmlns:w="` + packaging.NSWordprocessingML + `" xmlns:x="urn:opaque"><w:body>` + body + `<x:unknown x:type="x:Kind"/></w:body></w:document>`)
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, main)
		_, _ = q.AddPart("custom/opaque.bin", "application/octet-stream", []byte{0, 1, 99, 255})
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		if condition == "unknown settings extension" {
			_, _ = q.AddPart("word/settings.xml", packaging.ContentTypeSettings, []byte(`<w:settings xmlns:w="`+packaging.NSWordprocessingML+`" xmlns:x="urn:extension"><x:restriction enabled="1"/></w:settings>`))
			q.AddRelationship("word/document.xml", "settings.xml", packaging.RelTypeSettings)
		}
		if condition == "active document protection" {
			_, _ = q.AddPart("word/settings.xml", packaging.ContentTypeSettings, []byte(`<w:settings xmlns:w="`+packaging.NSWordprocessingML+`"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>`))
			q.AddRelationship("word/document.xml", "settings.xml", packaging.RelTypeSettings)
		}
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		session, err = document.OpenEditing(source, packaging.Limits{})
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		dir = ""
		out = ""
		target = nil
		failure = nil
		session = nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if dir != "" {
			return ctx, os.RemoveAll(dir)
		}
		return ctx, nil
	})
	sc.Step(`^a Word package containing a plain body text run and opaque content$`, func() error { return setup("") })
	sc.Step(`^a Word package with "([^"]+)"$`, setup)
	sc.Step(`^I replace the unique text "([^"]+)" with "([^"]+)" through a guarded target$`, func(old, replacement string) error {
		var err error
		target, err = session.FindOne(old)
		if err != nil {
			return err
		}
		return session.Replace(target, replacement)
	})
	save := func() error {
		var err error
		dir, err = os.MkdirTemp("", "word-edit-")
		if err != nil {
			return err
		}
		out = filepath.Join(dir, "output.docx")
		_, err = session.SaveAs(out)
		return err
	}
	sc.Step(`^save the edited Word package$`, save)
	sc.Step(`^only the matched XML character data differs$`, func() error {
		b, err := os.ReadFile(out)
		if err != nil {
			return err
		}
		parts, err := zipPayloads(b)
		if err != nil {
			return err
		}
		expected := strings.Replace(string(main), "old text", "new &amp; text", 1)
		if string(parts["word/document.xml"]) != expected {
			return fmt.Errorf("unexpected body edits")
		}
		return nil
	})
	sc.Step(`^every other package member retains its payload bytes$`, func() error {
		b, err := os.ReadFile(out)
		if err != nil {
			return err
		}
		before, err := zipPayloads(source)
		if err != nil {
			return err
		}
		after, err := zipPayloads(b)
		if err != nil {
			return err
		}
		if len(before) != len(after) {
			return fmt.Errorf("member set changed")
		}
		for n, v := range before {
			if n != "word/document.xml" && !bytes.Equal(v, after[n]) {
				return fmt.Errorf("member changed %s", n)
			}
		}
		return nil
	})
	sc.Step(`^reusing the consumed target returns a stale-target refusal$`, func() error {
		err := session.Replace(target, "another")
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("expected stale target, got %v", err)
		}
		return nil
	})
	sc.Step(`^I request a guarded correction of "([^"]+)"$`, func(old string) error {
		target, failure = session.FindOne(old)
		if failure == nil {
			failure = session.Replace(target, " new text ")
		}
		return nil
	})
	sc.Step(`^the correction returns a typed refusal$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected refusal, got %v", failure)
		}
		return nil
	})
	sc.Step(`^saving the Word session retains the original archive bytes$`, func() error {
		if err := save(); err != nil {
			return err
		}
		b, err := os.ReadFile(out)
		if err != nil {
			return err
		}
		if !bytes.Equal(b, source) {
			return fmt.Errorf("refusal changed archive")
		}
		return nil
	})
}
