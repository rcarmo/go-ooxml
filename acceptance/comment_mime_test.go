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

const paperCommentsExtendedType = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"

// Preservation characterisation only. Both spellings remain unclassified for
// native Office compatibility and full comment-thread semantics.
func commentMIMESteps(sc *godog.ScenarioContext) {
	var source, delivered []byte
	var mime string
	var session *document.EditSession
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		delivered = nil
		mime = ""
		session = nil
		return ctx, nil
	})
	sc.Step(`^a Word package with extended-comment MIME spelling "([^"]+)"$`, func(origin string) error {
		switch origin {
		case "pinned-go":
			mime = packaging.ContentTypeCommentsExtended
		case "specified-metadata":
			mime = paperCommentsExtendedType
		default:
			return fmt.Errorf("unknown source spelling")
		}
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body><w:p><w:r><w:t>ordinary body</w:t></w:r></w:p></w:body></w:document>`))
		_, _ = q.AddPart("word/comments.xml", packaging.ContentTypeComments, []byte(`<w:comments xmlns:w="`+packaging.NSWordprocessingML+`" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="Reviewer"><w:p w14:paraId="00000001"><w:r><w:t>Review note</w:t></w:r></w:p></w:comment></w:comments>`))
		_, _ = q.AddPart("word/commentsExtended.xml", mime, []byte(`<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="00000001" w15:done="0"/></w15:commentsEx>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("word/document.xml", "comments.xml", packaging.RelTypeComments)
		q.AddRelationship("word/document.xml", "commentsExtended.xml", packaging.RelTypeCommentsExtended)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = bytes.Clone(b.Bytes())
		var err error
		session, err = document.OpenEditing(source, packaging.Limits{})
		return err
	})
	sc.Step(`^I make an unrelated guarded body text correction$`, func() error {
		target, err := session.FindOne("ordinary body")
		if err != nil {
			return err
		}
		if err = session.Replace(target, "updated body"); err != nil {
			return err
		}
		dir, err := os.MkdirTemp("", "comment-mime-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.docx")
		if _, err = session.SaveAs(path); err != nil {
			return err
		}
		delivered, err = os.ReadFile(path)
		return err
	})
	sc.Step(`^the delivered package retains the exact extended-comment MIME spelling$`, func() error {
		q, err := packaging.OpenBytes(delivered)
		if err != nil {
			return err
		}
		defer q.Close()
		if q.GetContentType("word/commentsExtended.xml") != mime {
			return fmt.Errorf("MIME spelling changed")
		}
		before, err := zipPayloads(source)
		if err != nil {
			return err
		}
		after, err := zipPayloads(delivered)
		if err != nil {
			return err
		}
		if !bytes.Equal(before["[Content_Types].xml"], after["[Content_Types].xml"]) {
			return fmt.Errorf("content type registry bytes changed")
		}
		return nil
	})
	sc.Step(`^extended comments and relationship payloads remain byte-identical$`, func() error {
		before, err := zipPayloads(source)
		if err != nil {
			return err
		}
		after, err := zipPayloads(delivered)
		if err != nil {
			return err
		}
		if len(before) != len(after) {
			return fmt.Errorf("part set changed")
		}
		for name, b := range before {
			if name != "word/document.xml" && !bytes.Equal(b, after[name]) {
				return fmt.Errorf("unrelated part changed: %s", name)
			}
		}
		return nil
	})
}
