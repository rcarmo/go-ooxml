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
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

func slideSteps(sc *godog.ScenarioContext) {
	var source, slide []byte
	var s *presentation.EditSession
	var target *presentation.TextTarget
	var failure error
	var condition, dir string
	setup := func(c string) error {
		condition = c
		run := `<a:r><a:rPr b="1"/><a:t>old text</a:t></a:r>`
		if c == "repeated leaf text" {
			run += run
		}
		if c == "field-bearing paragraph" {
			run += `<a:fld id="{x}" type="slidenum"><a:t>1</a:t></a:fld>`
		}
		shape := `<p:sp><p:nvSpPr><p:cNvPr id="7" name="Title"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>` + run + `</a:p></p:txBody></p:sp>`
		if c == "duplicate shape IDs" {
			shape += shape
		}
		slide = []byte(`<p:sld xmlns:p="` + packaging.NSPresentationML + `" xmlns:a="` + packaging.NSDrawingML + `" xmlns:x="urn:opaque"><p:cSld><p:spTree>` + shape + `</p:spTree><x:unknown x:v='keep'/></p:cSld></p:sld>`)
		q := packaging.New()
		_, _ = q.AddPart("ppt/presentation.xml", packaging.ContentTypePresentation, []byte(`<p:presentation xmlns:p="`+packaging.NSPresentationML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`))
		_, _ = q.AddPart("ppt/slides/slide1.xml", packaging.ContentTypeSlide, slide)
		q.AddRelationship("", "ppt/presentation.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("ppt/presentation.xml", "slides/slide1.xml", packaging.RelTypeSlide)
		_, _ = q.AddPart("custom/opaque.bin", "application/octet-stream", []byte("unchanged"))
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = presentation.OpenEditing(source, packaging.Limits{})
		return err
	}
	save := func() ([]byte, error) {
		var err error
		dir, err = os.MkdirTemp("", "slide-edit-")
		if err != nil {
			return nil, err
		}
		out := filepath.Join(dir, "out.pptx")
		if _, err = s.SaveAs(out); err != nil {
			return nil, err
		}
		return os.ReadFile(out)
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		dir = ""
		failure = nil
		target = nil
		s = nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if dir != "" {
			return ctx, os.RemoveAll(dir)
		}
		return ctx, nil
	})
	sc.Step(`^a presentation with a plain text shape and opaque slide markup$`, func() error { return setup("") })
	sc.Step(`^a presentation text shape with "([^"]+)"$`, setup)
	sc.Step(`^I replace "([^"]+)" in shape (\d+) on "([^"]+)"$`, func(text string, id int, part string) error {
		var err error
		target, err = s.FindText(part, uint32(id), text)
		if err != nil {
			return err
		}
		return s.Replace(target, "new & text")
	})
	sc.Step(`^only the selected slide text bytes differ after delivery$`, func() error {
		data, err := save()
		if err != nil {
			return err
		}
		before, err := zipPayloads(source)
		if err != nil {
			return err
		}
		after, err := zipPayloads(data)
		if err != nil {
			return err
		}
		if len(before) != len(after) {
			return fmt.Errorf("member set changed")
		}
		for name, b := range before {
			want := b
			if name == "ppt/slides/slide1.xml" {
				want = []byte(strings.Replace(string(slide), "old text", "new &amp; text", 1))
			}
			if !bytes.Equal(want, after[name]) {
				return fmt.Errorf("changed bytes %s", name)
			}
		}
		return nil
	})
	sc.Step(`^the consumed presentation target refuses reuse$`, func() error {
		err := s.Replace(target, "again")
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("got %v", err)
		}
		return nil
	})
	sc.Step(`^I request a guarded presentation text correction$`, func() error {
		part := "ppt/slides/slide1.xml"
		if condition == "foreign slide part" {
			part = "ppt/slides/absent.xml"
		}
		target, failure = s.FindText(part, 7, "old text")
		if failure == nil {
			failure = s.Replace(target, "new")
		}
		return nil
	})
	sc.Step(`^the presentation operation returns a typed refusal$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("got %v", failure)
		}
		return nil
	})
	sc.Step(`^its delivered archive is byte-identical to the source$`, func() error {
		data, err := save()
		if err != nil {
			return err
		}
		if !bytes.Equal(data, source) {
			return fmt.Errorf("refusal changed source")
		}
		return nil
	})
}
