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

func storySteps(sc *godog.ScenarioContext) {
	var s *document.EditSession
	var source []byte
	var outline document.StoryOutline
	setup := func(nested bool) error {
		body := `<w:p><w:r><w:t>base </w:t></w:r><w:del w:id="1"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="2"><w:r><w:t>new</w:t></w:r></w:ins></w:p>`
		if nested {
			body = `<w:p><w:r><w:t>Outer</w:t><w:pict><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>`
		}
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body+`</w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		_, _ = q.AddPart("word/header1.xml", packaging.ContentTypeHeader, []byte(`<w:hdr xmlns:w="`+packaging.NSWordprocessingML+`"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`))
		q.AddRelationship("word/document.xml", "header1.xml", packaging.RelTypeHeader)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = document.OpenEditing(source, packaging.Limits{})
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		outline = document.StoryOutline{}
		return ctx, nil
	})
	sc.Step(`^a Word document with body revisions and a related header$`, func() error { return setup(false) })
	sc.Step(`^a Word document with a nested text box paragraph$`, func() error { return setup(true) })
	sc.Step(`^I inspect its "([^"]+)" story view$`, func(view string) error { var err error; outline, err = s.Outline(document.View(view)); return err })
	check := func(part, text string, box bool) error {
		n := 0
		for _, b := range outline.Blocks {
			if b.Part == part && b.Text == text && b.InTextBox == box {
				n++
			}
		}
		if n != 1 {
			return fmt.Errorf("expected one %s %q box=%v: %+v", part, text, box, outline)
		}
		return nil
	}
	sc.Step(`^the body paragraph reads "([^"]+)"$`, func(text string) error { return check("word/document.xml", text, false) })
	sc.Step(`^the related header paragraph reads "([^"]+)"$`, func(text string) error { return check("word/header1.xml", text, false) })
	sc.Step(`^the outer paragraph reads "([^"]+)"$`, func(text string) error { return check("word/document.xml", text, false) })
	sc.Step(`^one text-box paragraph reads "([^"]+)"$`, func(text string) error { return check("word/document.xml", text, true) })
	sc.Step(`^story inspection leaves the source archive unchanged$`, func() error {
		dir, err := os.MkdirTemp("", "story-view-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.docx")
		if _, err = s.SaveAs(path); err != nil {
			return err
		}
		got, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		if !bytes.Equal(source, got) {
			return fmt.Errorf("inspection changed archive")
		}
		return nil
	})
}
