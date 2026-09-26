package document

import (
	"bytes"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestSelectedBatchContracts(t *testing.T) {
	open := func() *EditSession {
		t.Helper()
		q := packaging.New()
		_, _ = q.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body><w:p><w:r><w:t xml:space="preserve">alpha beta alpha</w:t></w:r></w:p></w:body></w:document>`))
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		_ = q.WriteTo(&b)
		s, err := OpenEditing(b.Bytes(), packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		return s
	}
	t.Run("different snapshots same leaf", func(t *testing.T) {
		s := open()
		a, _ := s.FindText("alpha")
		b, _ := s.FindOne("beta")
		if err := s.ReplaceBatch([]TextReplacement{{a[0], "one"}, {b, "two"}, {a[1], "three"}}); err != nil {
			t.Fatal(err)
		}
		o, _ := s.Outline(CurrentView)
		if o.Blocks[0].Text != "one two three" {
			t.Fatal(o)
		}
	})
	t.Run("duplicate overlap invalid and no op", func(t *testing.T) {
		s := open()
		a, _ := s.FindText("alpha")
		whole, _ := s.FindOne("alpha beta")
		for _, changes := range [][]TextReplacement{{{a[0], "one"}, {a[0], "two"}}, {{a[0], "one"}, {whole, "two"}}, {{a[0], "one"}, {a[1], "\x00"}}} {
			if err := s.ReplaceBatch(changes); err == nil {
				t.Fatal("invalid batch accepted")
			}
			o, _ := s.Outline(CurrentView)
			if o.Blocks[0].Text != "alpha beta alpha" {
				t.Fatal("partial mutation")
			}
		}
		if err := s.ReplaceBatch([]TextReplacement{{a[0], "alpha"}, {a[1], "alpha"}}); err != nil {
			t.Fatal(err)
		}
		if err := s.Replace(a[0], "new"); err != nil {
			t.Fatal("no-op consumed target", err)
		}
	})
}
