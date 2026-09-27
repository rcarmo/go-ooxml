package presentation

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func notesWithParagraphs(t *testing.T, paragraphs string) *EditSession {
	t.Helper()
	s, _ := openNotesFixture(t)
	part := "ppt/notesSlides/notesSlide1.xml"
	target, err := s.FindNotes("ppt/slides/slide1.xml")
	if err != nil {
		t.Fatal(err)
	}
	old := target.paragraphs[0].Raw()
	b, h, err := s.pkg.Part(part)
	if err != nil {
		t.Fatal(err)
	}
	b = bytes.Replace(b, old, []byte(paragraphs), 1)
	if err = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: b}}); err != nil {
		t.Fatal(err)
	}
	return s
}
func TestMultilineNotesEdges(t *testing.T) {
	const slide = "ppt/slides/slide1.xml"
	const part = "ppt/notesSlides/notesSlide1.xml"
	t.Run("self-closing text fill", func(t *testing.T) {
		s := notesWithParagraphs(t, `<a:p><a:r><a:rPr b="1"/><a:t/></a:r></a:p>`)
		n, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(n, "filled"); err != nil {
			t.Fatal(err)
		}
		got, err := s.FindNotes(slide)
		if err != nil || got.Text() != "filled" {
			t.Fatal(err)
		}
	})
	t.Run("single leaf edge whitespace enables preservation", func(t *testing.T) {
		s := notesWithParagraphs(t, `<a:p><a:r><a:t xml:space="default">old</a:t></a:r></a:p>`)
		n, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(n, " leading "); err != nil {
			t.Fatal(err)
		}
		b, _, _ := s.pkg.Part(part)
		d, err := losslessxml.Parse(b)
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for _, e := range d.Elements() {
			text, _ := e.Text()
			if e.Name() == name(packaging.NSDrawingML, "t") && text == " leading " {
				for _, a := range e.Attributes() {
					if a.Name == (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) && a.Value == "preserve" {
						found = true
					}
				}
			}
		}
		if !found {
			t.Fatal("edge spaces not marked preserve")
		}
	})
	t.Run("first run no format never borrows second run", func(t *testing.T) {
		s := notesWithParagraphs(t, `<a:p><a:pPr algn="ctr"/><a:r><a:t>one</a:t></a:r><a:r><a:rPr b="1"/><a:t>two</a:t></a:r></a:p>`)
		n, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		before, _, _ := s.pkg.Part(part)
		if err = s.ReplaceNotes(n, n.Text()); err != nil {
			t.Fatal(err)
		}
		after, _, _ := s.pkg.Part(part)
		if !bytes.Equal(before, after) {
			t.Fatal("multirun no-op reserialized")
		}
		if err = s.ReplaceNotes(n, "\n A&B \n雪\n"); err != nil {
			t.Fatal(err)
		}
		newTarget, err := s.FindNotes(slide)
		if err != nil || newTarget.Text() != "\n A&B \n雪\n" {
			t.Fatal(err)
		}
		if len(newTarget.paragraphs) != 4 {
			t.Fatal("boundary empty paragraphs lost")
		}
		b, _, _ := s.pkg.Part(part)
		if bytes.Contains(b, []byte(`b="1"`)) {
			t.Fatal("later run formatting borrowed")
		}
	})
	t.Run("template shape and child values survive", func(t *testing.T) {
		s := notesWithParagraphs(t, `<a:p><a:pPr algn="ctr"><a:buChar char="•"/><a:defRPr sz="1200"/></a:pPr><a:r><a:rPr b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="F&amp;F"/></a:rPr><a:t>old</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`)
		n, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(n, "a\nb"); err != nil {
			t.Fatal(err)
		}
		b, _, _ := s.pkg.Part(part)
		for _, fragment := range []string{`val="112233"`, `typeface="F&amp;F"`, `lang="en-US"`, `char="•"`} {
			if strings.Count(string(b), fragment) != 2 {
				t.Fatal("template property lost", fragment)
			}
		}
	})
}
