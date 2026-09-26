package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func imageReplaceSteps(sc *godog.ScenarioContext) {
	var s *spreadsheet.EditSession
	var source, output, replacement []byte
	var target *spreadsheet.ImageTarget
	var failure error
	save := func() error {
		dir, err := os.MkdirTemp("", "picture-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.xlsx")
		if _, err = s.SaveAs(path); err != nil {
			return err
		}
		output, err = os.ReadFile(path)
		return err
	}
	setup := func(condition string) error {
		q := packaging.New()
		ns := packaging.NSSpreadsheetML
		rel := packaging.NSDocumentRelationships
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+ns+`" xmlns:r="`+rel+`"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`))
		protection := ""
		if condition == "protected worksheet" {
			protection = `<sheetProtection sheet="1"/>`
		}
		_, _ = q.AddPart("xl/sheet.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+ns+`" xmlns:r="`+rel+`"><sheetData/>`+protection+`<drawing r:id="rId1"/></worksheet>`))
		pic := func(id int, relationship string) string {
			return fmt.Sprintf(`<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="%d" name="Picture %d"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="%s"/><a:srcRect l="100"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`, id, id, relationship)
		}
		second := "rId2"
		if condition == "shared relationship ID" {
			second = "rId1"
		}
		drawing := `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="` + packaging.NSDrawingML + `" xmlns:r="` + rel + `">` + pic(1, "rId1") + pic(2, second) + `</xdr:wsDr>`
		_, _ = q.AddPart("xl/drawings/drawing.xml", packaging.ContentTypeDrawing, []byte(drawing))
		var pngBytes bytes.Buffer
		im := image.NewRGBA(image.Rect(0, 0, 1, 1))
		im.Set(0, 0, color.RGBA{R: 255, A: 255})
		if err := png.Encode(&pngBytes, im); err != nil {
			return err
		}
		_, _ = q.AddPart("xl/media/shared.png", packaging.ContentTypePNG, bytes.Clone(pngBytes.Bytes()))
		pngBytes.Reset()
		im.Set(0, 0, color.RGBA{B: 255, A: 255})
		if err := png.Encode(&pngBytes, im); err != nil {
			return err
		}
		replacement = bytes.Clone(pngBytes.Bytes())
		if condition == "malformed image data" {
			replacement = []byte("not a PNG")
		}
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "sheet.xml", packaging.RelTypeWorksheet)
		q.AddRelationship("xl/sheet.xml", "drawings/drawing.xml", packaging.RelTypeDrawing)
		if condition == "external picture" {
			q.AddRelationshipWithTargetMode("xl/drawings/drawing.xml", "https://example.invalid/p.png", packaging.RelTypeImage, packaging.TargetModeExternal)
		} else {
			q.AddRelationship("xl/drawings/drawing.xml", "../media/shared.png", packaging.RelTypeImage)
		}
		q.AddRelationship("xl/drawings/drawing.xml", "../media/shared.png", packaging.RelTypeImage)
		if condition == "shared drawing part" {
			_, _ = q.AddPart("custom/owner.xml", packaging.ContentTypeXML, []byte(`<r/>`))
			q.AddRelationship("custom/owner.xml", "../xl/drawings/drawing.xml", packaging.RelTypeDrawing)
		}
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = spreadsheet.OpenEditing(source, packaging.Limits{})
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		target = nil
		source = nil
		output = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a worksheet with two pictures sharing old image bytes$`, func() error { return setup("") })
	sc.Step(`^an image workbook with "([^"]+)"$`, setup)
	sc.Step(`^I replace picture one with a valid PNG$`, func() error {
		var err error
		target, err = s.FindImage("S", 1)
		if err != nil {
			return err
		}
		if err = s.ReplaceImage(target, replacement); err != nil {
			return err
		}
		return save()
	})
	sc.Step(`^drawing geometry and the other picture remain byte-identical$`, func() error {
		before, _ := zipPayloads(source)
		after, err := zipPayloads(output)
		if err != nil {
			return err
		}
		for name, b := range before {
			if name == "[Content_Types].xml" || name == "xl/drawings/_rels/drawing.xml.rels" {
				continue
			}
			if !bytes.Equal(b, after[name]) {
				return fmt.Errorf("unexpected payload change %s", name)
			}
		}
		return nil
	})
	sc.Step(`^the selected image uses fresh media while original media is retained$`, func() error {
		q, err := packaging.OpenPreserved(output, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		found := false
		for _, e := range g.Edges {
			if e.Source != "xl/drawings/drawing.xml" {
				continue
			}
			if e.ID == "rId1" {
				if e.ResolvedPart == "xl/media/shared.png" || !strings.HasPrefix(e.ResolvedPart, "xl/media/") {
					return fmt.Errorf("no fresh media")
				}
				data, _, err := q.Part(e.ResolvedPart)
				if err != nil || !bytes.Equal(data, replacement) {
					return fmt.Errorf("replacement payload differs")
				}
				found = true
			}
			if e.ID == "rId2" && e.ResolvedPart != "xl/media/shared.png" {
				return fmt.Errorf("shared consumer redirected")
			}
		}
		if !found {
			return fmt.Errorf("selected relationship missing")
		}
		return nil
	})
	sc.Step(`^the consumed image target refuses reuse$`, func() error {
		err := s.ReplaceImage(target, replacement)
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("stale target accepted: %v", err)
		}
		return nil
	})
	sc.Step(`^I attempt to replace its selected picture$`, func() error {
		target, failure = s.FindImage("S", 1)
		if failure == nil {
			failure = s.ReplaceImage(target, replacement)
		}
		return nil
	})
	sc.Step(`^image replacement refuses without package mutation$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected image refusal %v", failure)
		}
		if err := save(); err != nil {
			return err
		}
		if !bytes.Equal(source, output) {
			return fmt.Errorf("refusal mutated archive")
		}
		return nil
	})
}
