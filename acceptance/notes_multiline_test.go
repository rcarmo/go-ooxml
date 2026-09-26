package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

func notesMultilineSteps(sc *godog.ScenarioContext) {
	var s *presentation.EditSession
	var source, out, originalNotes []byte
	var condition string
	var failure error
	var target *presentation.NotesTarget
	var cleared []byte
	const part = "ppt/notesSlides/notes.xml"
	const slide = "ppt/slides/s.xml"
	const wanted = "First line\nSecond line\n\nFourth line"
	const before = `<a:bodyPr/><a:lstStyle/><!--before paragraphs-->`
	const after = `<!--after paragraphs-->`
	save := func() error {
		dir, err := os.MkdirTemp("", "notes-lines-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.pptx")
		if _, err = s.SaveAs(path); err != nil {
			return err
		}
		out, err = os.ReadFile(path)
		return err
	}
	body := func(data []byte) (*losslessxml.Document, []losslessxml.Element, error) {
		parts, err := zipPayloads(data)
		if err != nil {
			return nil, nil, err
		}
		d, err := losslessxml.Parse(parts[part])
		if err != nil {
			return nil, nil, err
		}
		var paras []losslessxml.Element
		for _, e := range d.Elements() {
			if e.Name() == (xml.Name{Space: packaging.NSDrawingML, Local: "p"}) {
				paras = append(paras, e)
			}
		}
		return d, paras, nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		out = nil
		cleared = nil
		condition = ""
		failure = nil
		target = nil
		return ctx, nil
	})
	sc.Step(`^multiline notes condition "([^"]+)"$`, func(c string) error {
		condition = c
		q := packaging.New()
		p, a, r := packaging.NSPresentationML, packaging.NSDrawingML, packaging.NSDocumentRelationships
		_, _ = q.AddPart("ppt/presentation.xml", packaging.ContentTypePresentation, []byte(`<p:presentation xmlns:p="`+p+`" xmlns:r="`+r+`"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`))
		_, _ = q.AddPart(slide, packaging.ContentTypeSlide, []byte(`<p:sld xmlns:p="`+p+`"><p:cSld><p:spTree/></p:cSld></p:sld>`))
		q.AddRelationship("", "ppt/presentation.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("ppt/presentation.xml", "slides/s.xml", packaging.RelTypeSlide)
		q.AddRelationship(slide, "../notesSlides/notes.xml", packaging.RelTypeNotesSlide)
		paragraphs := `<a:p><a:pPr algn="ctr" lvl="1"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr><a:r><a:rPr b="1" sz="1400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Source Sans"/></a:rPr><a:t>Original</a:t></a:r><a:r><a:rPr i="1"/><a:t> text</a:t></a:r></a:p><a:p><a:pPr algn="r"/><a:r><a:t>Second paragraph</a:t></a:r></a:p>`
		switch c {
		case "multiple runs and paragraphs", "invalid character after newline":
		case "empty existing paragraph":
			paragraphs = `<a:p><a:pPr algn="ctr" lvl="1"/></a:p>`
		case "prefixed template namespaces":
			paragraphs = strings.Replace(paragraphs, `<a:rPr b="1"`, `<a:rPr xmlns:z="`+a+`" b="1"`, 1)
			paragraphs = strings.ReplaceAll(paragraphs, "a:solidFill", "z:solidFill")
			paragraphs = strings.ReplaceAll(paragraphs, "a:srgbClr", "z:srgbClr")
		case "hyperlink in later run":
			paragraphs = strings.Replace(paragraphs, `<a:rPr i="1"/>`, `<a:rPr i="1"><a:hlinkClick xmlns:r="`+r+`" r:id="rId2"/></a:rPr>`, 1)
		case "unknown paragraph property":
			paragraphs = strings.Replace(paragraphs, `<a:lnSpc>`, `<a:unproved/><a:lnSpc>`, 1)
		default:
			return fmt.Errorf("unknown condition")
		}
		originalNotes = []byte(`<p:notes xmlns:p="` + p + `" xmlns:a="` + a + `"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody>` + before + paragraphs + after + `</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
		_, _ = q.AddPart(part, packaging.ContentTypeNotesSlide, originalNotes)
		_, _ = q.AddPart("opaque.bin", "application/octet-stream", []byte("opaque"))
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = presentation.OpenEditing(source, packaging.Limits{})
		return err
	})
	sc.Step(`^I replace the existing notes with four lines including an empty line$`, func() error {
		var err error
		target, err = s.FindNotes(slide)
		if err != nil {
			return err
		}
		expected := "Original text\nSecond paragraph"
		if condition == "empty existing paragraph" {
			expected = ""
		}
		if target.Text() != expected {
			return fmt.Errorf("readback got%q", target.Text())
		}
		if err = s.ReplaceNotes(target, wanted); err != nil {
			return err
		}
		return save()
	})
	sc.Step(`^reopened notes contain exactly the four requested lines$`, func() error {
		reopened, err := presentation.OpenEditing(out, packaging.Limits{})
		if err != nil {
			return err
		}
		n, err := reopened.FindNotes(slide)
		if err != nil || n.Text() != wanted {
			return fmt.Errorf("readback %v %v", n, err)
		}
		_, ps, err := body(out)
		if err != nil || len(ps) != 4 {
			return fmt.Errorf("paragraph inventory %d %v", len(ps), err)
		}
		return nil
	})
	sc.Step(`^each new paragraph keeps the first paragraph formatting$`, func() error {
		d, ps, err := body(out)
		if err != nil {
			return err
		}
		for _, p := range ps {
			found := false
			for _, e := range d.Elements() {
				if parent, ok := e.Parent(); ok && parent == p && e.Name() == (xml.Name{Space: packaging.NSDrawingML, Local: "pPr"}) {
					algn, lvl := "", ""
					for _, a := range e.Attributes() {
						if a.Name.Local == "algn" {
							algn = a.Value
						}
						if a.Name.Local == "lvl" {
							lvl = a.Value
						}
					}
					found = algn == "ctr" && lvl == "1"
				}
			}
			if !found {
				return fmt.Errorf("paragraph template missing")
			}
		}
		return nil
	})
	sc.Step(`^each nonempty line keeps the first run formatting without formatting from later runs$`, func() error {
		d, ps, err := body(out)
		if err != nil {
			return err
		}
		for i, p := range ps {
			runs := 0
			for _, e := range d.Elements() {
				if parent, ok := e.Parent(); ok && parent == p && e.Name() == (xml.Name{Space: packaging.NSDrawingML, Local: "r"}) {
					runs++
					if condition == "empty existing paragraph" {
						continue
					}
					found := false
					for _, prop := range d.Elements() {
						if parent, ok := prop.Parent(); ok && parent == e && prop.Name().Local == "rPr" {
							flags := map[string]string{}
							for _, a := range prop.Attributes() {
								flags[a.Name.Local] = a.Value
							}
							found = flags["b"] == "1" && flags["sz"] == "1400" && flags["i"] == ""
						}
					}
					if !found {
						return fmt.Errorf("run template wrong")
					}
				}
			}
			expected := 1
			if i == 2 {
				expected = 0
			}
			if runs != expected {
				return fmt.Errorf("line%d runs%d", i, runs)
			}
		}
		return nil
	})
	sc.Step(`^all non-body notes XML and other package payloads stay unchanged$`, func() error {
		a, _ := zipPayloads(source)
		b, err := zipPayloads(out)
		if err != nil {
			return err
		}
		if len(a) != len(b) {
			return fmt.Errorf("parts changed")
		}
		for p, data := range a {
			if p != part && !bytes.Equal(data, b[p]) {
				return fmt.Errorf("changed part %s", p)
			}
		}
		old, new := a[part], b[part]
		prefix := bytes.Index(old, []byte(before)) + len(before)
		suffix := bytes.Index(old, []byte(after))
		if !bytes.HasPrefix(new, old[:prefix]) || !bytes.HasSuffix(new, old[suffix:]) {
			return fmt.Errorf("outside paragraphs changed")
		}
		return nil
	})
	sc.Step(`^I clear and refill the existing notes body$`, func() error {
		n, err := s.FindNotes(slide)
		if err != nil {
			return err
		}
		if err = s.ReplaceNotes(n, ""); err != nil {
			return err
		}
		if err = save(); err != nil {
			return err
		}
		cleared = bytes.Clone(out)
		n, err = s.FindNotes(slide)
		if err != nil {
			return err
		}
		if err = s.ReplaceNotes(n, "refilled"); err != nil {
			return err
		}
		return save()
	})
	sc.Step(`^the cleared body is one empty paragraph and refill has the requested text$`, func() error {
		d, ps, err := body(cleared)
		if err != nil || len(ps) != 1 {
			return fmt.Errorf("cleared paragraphs")
		}
		for _, e := range d.Elements() {
			if e.Name() == (xml.Name{Space: packaging.NSDrawingML, Local: "r"}) {
				return fmt.Errorf("cleared paragraph has run")
			}
		}
		r, err := presentation.OpenEditing(out, packaging.Limits{})
		if err != nil {
			return err
		}
		n, err := r.FindNotes(slide)
		if err != nil || n.Text() != "refilled" {
			return fmt.Errorf("refill failed %v", err)
		}
		return nil
	})
	sc.Step(`^I attempt an unsupported multiline notes edit$`, func() error {
		target, failure = s.FindNotes(slide)
		if failure == nil {
			replacement := wanted
			if condition == "invalid character after newline" {
				replacement = "good\ninvalid\x00"
			}
			failure = s.ReplaceNotes(target, replacement)
		}
		return nil
	})
	sc.Step(`^the notes snapshot and package remain unchanged$`, func() error {
		var refusal *packaging.Refusal
		if !errors.As(failure, &refusal) {
			return fmt.Errorf("missing refusal %v", failure)
		}
		if err := save(); err != nil {
			return err
		}
		if !bytes.Equal(source, out) {
			return fmt.Errorf("failure removed original paragraphs")
		}
		return nil
	})
}
