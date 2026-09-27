package document

import (
	"bytes"
	"errors"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestEditingTargetAndGuardBatch(t *testing.T) {
	open := func(body string) *EditSession {
		t.Helper()
		p := packaging.New()
		_, _ = p.AddPart("word/document.xml", packaging.ContentTypeWordDocument, []byte(`<w:document xmlns:w="`+packaging.NSWordprocessingML+`"><w:body>`+body+`</w:body></w:document>`))
		p.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		var b bytes.Buffer
		if err := p.WriteTo(&b); err != nil {
			t.Fatal(err)
		}
		s, err := OpenEditing(b.Bytes(), packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		return s
	}
	plain := `<w:p><w:r><w:t xml:space="preserve">one</w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>`
	t.Run("no-op reuse foreign and stale", func(t *testing.T) {
		s := open(plain)
		a, err := s.FindOne("one")
		if err != nil {
			t.Fatal(err)
		}
		b, _ := s.FindOne("two")
		if err = s.Replace(a, "one"); err != nil {
			t.Fatal(err)
		}
		other := open(plain)
		if err = other.Replace(a, "new"); err == nil {
			t.Fatal("foreign accepted")
		}
		if err = s.Replace(a, "new"); err != nil {
			t.Fatal(err)
		}
		err = s.Replace(b, "newer")
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			t.Fatalf("got %v", err)
		}
		fresh, err := s.FindOne("two")
		if err != nil {
			t.Fatal(err)
		}
		if err = s.Replace(fresh, "next"); err != nil {
			t.Fatal(err)
		}
	})
	for _, body := range []string{`<w:p><w:hyperlink><w:r><w:t>one</w:t></w:r></w:hyperlink></w:p>`, `<w:p><w:r><w:t>one</w:t><w:br/></w:r></w:p>`, `<w:sdt><w:p><w:r><w:t>one</w:t></w:r></w:p></w:sdt>`} {
		t.Run(body, func(t *testing.T) {
			s := open(body)
			target, err := s.FindOne("one")
			if err != nil {
				t.Fatal(err)
			}
			if err = s.Replace(target, "new"); err == nil {
				t.Fatal("unsafe wrapper accepted")
			}
			if text, _ := target.element.Text(); text != "one" {
				t.Fatal("refusal mutated target")
			}
		})
	}
	t.Run("escaped invalid characters refusal reusable", func(t *testing.T) {
		s := open(plain)
		a, _ := s.FindOne("one")
		if err := s.Replace(a, "\x00"); err == nil {
			t.Fatal("invalid char accepted")
		}
		if err := s.Replace(a, "😀 & <ok>"); err != nil {
			t.Fatal(err)
		}
	})
}
