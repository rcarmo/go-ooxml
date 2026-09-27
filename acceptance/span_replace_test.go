package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func replaceSpanSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var source, main, output []byte
	var target *document.TextTarget
	var failure error
	var first, second, wantFirst, wantSecond string
	esc := func(s string) string { var b bytes.Buffer; _ = xml.EscapeText(&b, []byte(s)); return b.String() }
	run := func(text, prop string) string {
		return `<w:r><w:rPr><w:` + prop + `/></w:rPr><w:t xml:space="preserve">` + esc(text) + `</w:t></w:r>`
	}
	makeXML := func(a, b string, wrapper bool) []byte {
		body := run(a, "b") + run(b, "i")
		if wrapper {
			body = run(a, "b") + `<w:hyperlink>` + run(b, "i") + `</w:hyperlink>`
		}
		return []byte(`<w:document xmlns:w="` + packaging.NSWordprocessingML + `"><w:body><!--opaque--><w:p>` + body + `</w:p></w:body></w:document>`)
	}
	setup := func(a, b string, wrapper bool) error {
		first, second = a, b
		main = makeXML(a, b, wrapper)
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, main)
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var buf bytes.Buffer
		if err := q.WriteTo(&buf); err != nil {
			return err
		}
		source = buf.Bytes()
		var err error
		s, err = document.OpenEditing(source, packaging.Limits{})
		return err
	}
	save := func() ([]byte, error) {
		dir, err := os.MkdirTemp("", "span-replace-")
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
	unchanged := func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected refusal, got %v", failure)
		}
		data, err := save()
		if err != nil {
			return err
		}
		if !bytes.Equal(data, source) {
			return fmt.Errorf("refusal changed archive")
		}
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		output = nil
		target = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^editable Word runs "([^"]+)" and "([^"]+)"$`, func(a, b string) error { return setup(a, b, false) })
	sc.Step(`^text split across a plain run and a hyperlink wrapper$`, func() error { return setup("Al", "pha", true) })
	sc.Step(`^I replace the exact span "([^"]+)" with "([^"]+)"$`, func(needle, replacement string) error {
		var err error
		target, err = s.FindOne(needle)
		if err != nil {
			return err
		}
		if err = s.Replace(target, replacement); err != nil {
			return err
		}
		data, err := save()
		if err != nil {
			return err
		}
		parts, err := zipPayloads(data)
		if err != nil {
			return err
		}
		output = parts["word/document.xml"]
		return nil
	})
	sc.Step(`^the resulting run texts are "([^"]+)" and "([^"]+)"$`, func(a, b string) error { wantFirst, wantSecond = a, b; return nil })
	sc.Step(`^all run properties and surrounding XML are byte-preserved$`, func() error {
		want := makeXML(wantFirst, wantSecond, false)
		if !bytes.Equal(output, want) {
			return fmt.Errorf("unexpected XML\ngot %s\nwant %s (input %q/%q)", output, want, first, second)
		}
		return nil
	})
	sc.Step(`^I attempt to replace "([^"]+)" with "([^"]+)"$`, func(needle, replacement string) error {
		var err error
		target, err = s.FindOne(needle)
		if err != nil {
			return err
		}
		failure = s.Replace(target, replacement)
		return nil
	})
	sc.Step(`^ambiguous alignment refuses with the original archive intact$`, unchanged)
	sc.Step(`^wrapper-crossing replacement refuses with the original archive intact$`, unchanged)
	sc.Step(`^the same span remains reusable for a different replacement$`, func() error { return s.Replace(target, "Notice") })
}
