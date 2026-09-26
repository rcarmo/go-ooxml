package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func whitespaceSteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var output []byte
	setup := func(text string) error {
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`" xmlns:x="urn:opaque"><w:body><w:p><w:r><w:t x:keep = 'yes'>`+text+`</w:t></w:r></w:p></w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		var err error
		s, err = document.OpenEditing(b.Bytes(), packaging.Limits{})
		return err
	}
	save := func() error {
		dir, err := os.MkdirTemp("", "word-space-")
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
		parts, err := zipPayloads(b)
		if err != nil {
			return err
		}
		output = parts["word/document.xml"]
		return nil
	}
	check := func(want string) error {
		d, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		for _, e := range d.Elements() {
			if e.Name() == (xml.Name{Space: packaging.NSWordprocessingML, Local: "t"}) {
				text, _ := e.Text()
				if text != want {
					return fmt.Errorf("got %q expected %q", text, want)
				}
				count := 0
				for _, a := range e.Attributes() {
					if a.Name == (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) && a.Value == "preserve" {
						count++
					}
				}
				if count != 1 {
					return fmt.Errorf("expected one preservation attribute")
				}
				return nil
			}
		}
		return fmt.Errorf("text leaf absent")
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		output = nil
		return ctx, nil
	})
	sc.Step(`^a Word text leaf without XML space preservation$`, func() error { return setup("old") })
	sc.Step(`^two Word matches in a leaf without XML space preservation$`, func() error { return setup("red red") })
	sc.Step(`^I replace its text with leading and trailing spaces$`, func() error {
		target, err := s.FindOne("old")
		if err != nil {
			return err
		}
		if err = s.Replace(target, " new "); err != nil {
			return err
		}
		return save()
	})
	sc.Step(`^the changed leaf has XML space preserve and exact new text$`, func() error { return check(" new ") })
	sc.Step(`^its unrelated attribute bytes remain unchanged$`, func() error {
		if !bytes.Contains(output, []byte(`x:keep = 'yes'`)) {
			return fmt.Errorf("opaque attribute rewritten")
		}
		return nil
	})
	sc.Step(`^a selected batch adds significant spaces to both matches$`, func() error {
		spans, err := s.FindText("red")
		if err != nil {
			return err
		}
		changes := []document.TextReplacement{}
		for _, span := range spans {
			changes = append(changes, document.TextReplacement{Target: span, Text: " blue "})
		}
		if err = s.ReplaceBatch(changes); err != nil {
			return err
		}
		return save()
	})
	sc.Step(`^the saved leaf has one XML space preserve attribute$`, func() error {
		if bytes.Count(output, []byte(`xml:space="preserve"`)) != 1 {
			return fmt.Errorf("wrong preservation attribute count")
		}
		return nil
	})
	sc.Step(`^both replacements retain their significant spaces$`, func() error { return check(" blue   blue ") })
}
