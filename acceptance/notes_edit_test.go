package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

func notesEditSteps(sc *godog.ScenarioContext) {
	var s *presentation.EditSession
	var target *presentation.NotesTarget
	var source, out []byte
	var condition, oldText, newText string
	var failure error
	const slide = "ppt/slides/slide1.xml"
	save := func() error {
		dir, err := os.MkdirTemp("", "notes-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		p := filepath.Join(dir, "out.pptx")
		if _, err = s.SaveAs(p); err != nil {
			return err
		}
		out, err = os.ReadFile(p)
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		target = nil
		source = nil
		out = nil
		failure = nil
		condition = ""
		oldText = "Original notes"
		newText = "Updated notes"
		return ctx, nil
	})
	sc.Step(`^a presentation with notes condition "([^"]+)"$`, func(c string) error {
		condition = c
		q := packaging.New()
		pns, ans, rns := packaging.NSPresentationML, packaging.NSDrawingML, packaging.NSDocumentRelationships
		_, _ = q.AddPart("ppt/presentation.xml", packaging.ContentTypePresentation, []byte(`<p:presentation xmlns:p="`+pns+`" xmlns:r="`+rns+`"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`))
		_, _ = q.AddPart(slide, packaging.ContentTypeSlide, []byte(`<p:sld xmlns:p="`+pns+`"><p:cSld><p:spTree/></p:cSld></p:sld>`))
		q.AddRelationship("", "ppt/presentation.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("ppt/presentation.xml", "slides/slide1.xml", packaging.RelTypeSlide)
		_, _ = q.AddPart("opaque.bin", "application/octet-stream", []byte("sentinel"))
		if c != "absent notes" {
			if c == "empty text leaf" || c == "self-closing text leaf" {
				oldText = ""
			}
			body := func(id string) string {
				nv := `<p:cNvSpPr/>`
				if c == "grouping-only lock" {
					nv = `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`
				}
				if c == "malformed grouping lock" {
					nv = `<p:cNvSpPr><a:spLocks noGrp="maybe"/></p:cNvSpPr>`
				}
				if c == "locked body" {
					nv = `<p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr>`
				}
				run := `<a:r><a:rPr b="1"/><a:t>` + oldText + `</a:t></a:r>`
				if c == "self-closing text leaf" {
					run = `<a:r><a:rPr b="1"/><a:t/></a:r>`
				}
				if c == "field in body" {
					run = `<a:fld id="field"><a:t>` + oldText + `</a:t></a:fld>`
				}
				return `<p:sp><p:nvSpPr><p:cNvPr id="` + id + `" name="Notes"/>` + nv + `<p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>` + run + `</a:p></p:txBody></p:sp>`
			}
			bodies := body("2")
			if c == "duplicate body placeholder" {
				bodies += body("3")
			}
			other := `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="opaque"><a:t>7</a:t></a:fld></a:p></p:txBody></p:sp>`
			_, _ = q.AddPart("ppt/notesSlides/notes1.xml", packaging.ContentTypeNotesSlide, []byte(`<p:notes xmlns:p="`+pns+`" xmlns:a="`+ans+`"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>`+bodies+other+`</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`))
			q.AddRelationship(slide, "../notesSlides/notes1.xml", packaging.RelTypeNotesSlide)
			q.AddRelationship("ppt/notesSlides/notes1.xml", "../slides/slide1.xml", packaging.RelTypeSlide)
			if c == "shared notes part" {
				_, _ = q.AddPart("custom/owner.xml", packaging.ContentTypeXML, []byte(`<r/>`))
				q.AddRelationship("custom/owner.xml", "../ppt/notesSlides/notes1.xml", packaging.RelTypeNotesSlide)
			}
		}
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = presentation.OpenEditing(source, packaging.Limits{})
		return err
	})
	sc.Step(`^I find the existing notes body$`, func() error { var err error; target, err = s.FindNotes(slide); return err })
	sc.Step(`^the notes text is "([^"]*)"$`, func(want string) error {
		if target.Text() != want {
			return fmt.Errorf("got notes %q want %q", target.Text(), want)
		}
		return nil
	})
	sc.Step(`^I replace that notes body with "([^"]*)"$`, func(text string) error {
		newText = text
		if err := s.ReplaceNotes(target, text); err != nil {
			return err
		}
		return save()
	})
	sc.Step(`^only the notes body text bytes change after delivery$`, func() error {
		before, _ := zipPayloads(source)
		after, err := zipPayloads(out)
		if err != nil {
			return err
		}
		if len(before) != len(after) {
			return fmt.Errorf("notes part set changed")
		}
		for p, b := range before {
			want := b
			if p == "ppt/notesSlides/notes1.xml" {
				from := []byte(`<a:t>` + oldText + `</a:t>`)
				if condition == "self-closing text leaf" {
					from = []byte(`<a:t/>`)
				}
				want = bytes.Replace(b, from, []byte(`<a:t>`+newText+`</a:t>`), 1)
			}
			if !bytes.Equal(want, after[p]) {
				return fmt.Errorf("unexpected bytes %s", p)
			}
		}
		reopened, err := presentation.OpenEditing(out, packaging.Limits{})
		if err != nil {
			return err
		}
		body, err := reopened.FindNotes(slide)
		if err != nil || body.Text() != newText {
			return fmt.Errorf("notes reopen %v", err)
		}
		return nil
	})
	sc.Step(`^the old notes target refuses reuse$`, func() error {
		err := s.ReplaceNotes(target, "again")
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("old target accepted %v", err)
		}
		return nil
	})
	sc.Step(`^I attempt a guarded notes replacement$`, func() error {
		target, failure = s.FindNotes(slide)
		if failure != nil {
			return nil
		}
		if condition == "bad replacement character" {
			newText = "bad\x00"
		}
		if condition == "tab replacement" {
			newText = "two\tcolumns"
		}
		failure = s.ReplaceNotes(target, newText)
		return nil
	})
	sc.Step(`^notes replacement refuses and the archive is unchanged$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected typed refusal %v", failure)
		}
		if err := save(); err != nil {
			return err
		}
		if !bytes.Equal(source, out) {
			return fmt.Errorf("notes refusal changed archive")
		}
		return nil
	})
}
