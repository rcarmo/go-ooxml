package acceptance

import (
	"bytes"
	"context"
	"errors"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func officeLinkSteps(sc *godog.ScenarioContext) {
	var source, original []byte
	var format string
	var failure error
	var found bool
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		original = nil
		format = ""
		failure = nil
		found = false
		return ctx, nil
	})
	sc.Step(`^an Office link fixture for "([^"]+)" with "([^"]+)"$`, func(kind, condition string) error {
		format = kind
		q := packaging.New()
		rootAttrs := ` xmlns:r="` + packaging.NSDocumentRelationships + `"`
		attribute := `r:id="rId1"`
		switch condition {
		case "alias prefix":
			rootAttrs = ` xmlns:link="` + packaging.NSDocumentRelationships + `"`
			attribute = `link:id="rId1"`
		case "local correct binding":
			rootAttrs = ` xmlns:link="urn:foreign"`
			attribute = `xmlns:link="` + packaging.NSDocumentRelationships + `" link:id="rId1"`
		case "foreign r alongside correct link":
			rootAttrs = ` xmlns:r="urn:foreign" xmlns:link="` + packaging.NSDocumentRelationships + `"`
			attribute = `r:id="missing" link:id="rId1"`
		case "locally shadowed wrong URI":
			attribute = `xmlns:r="urn:foreign" r:id="rId1"`
		case "plain id with default namespace":
			attribute = `xmlns="` + packaging.NSDocumentRelationships + `" id="rId1"`
		case "foreign relationship type suffix":
		default:
			return fmt.Errorf("unknown condition")
		}
		switch kind {
		case "xlsx":
			ns := packaging.NSSpreadsheetML
			_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<s:workbook xmlns:s="`+ns+`"`+rootAttrs+`><s:sheets><s:sheet name="Sheet" sheetId="1" `+attribute+`/></s:sheets></s:workbook>`))
			_, _ = q.AddPart("xl/sheet.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+ns+`"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>`))
			q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
			typ := packaging.RelTypeWorksheet
			if condition == "foreign relationship type suffix" {
				typ = "urn:foreign/worksheet"
			}
			q.AddRelationship("xl/workbook.xml", "sheet.xml", typ)
		case "pptx":
			if condition == "plain id with default namespace" {
				attribute = `xmlns="` + packaging.NSDocumentRelationships + `"` /* existing numeric id must not be read as an r:id */
			}
			pns := packaging.NSPresentationML
			_, _ = q.AddPart("ppt/presentation.xml", packaging.ContentTypePresentation, []byte(`<p:presentation xmlns:p="`+pns+`"`+rootAttrs+`><p:sldIdLst><p:sldId id="256" `+attribute+`/></p:sldIdLst></p:presentation>`))
			_, _ = q.AddPart("ppt/slides/slide1.xml", packaging.ContentTypeSlide, []byte(`<p:sld xmlns:p="`+pns+`" xmlns:a="`+packaging.NSDrawingML+`"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`))
			q.AddRelationship("", "ppt/presentation.xml", packaging.RelTypeOfficeDocument)
			typ := packaging.RelTypeSlide
			if condition == "foreign relationship type suffix" {
				typ = "urn:foreign/slide"
			}
			q.AddRelationship("ppt/presentation.xml", "slides/slide1.xml", typ)
		default:
			return fmt.Errorf("unknown format")
		}
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		original = bytes.Clone(source)
		return nil
	})
	sc.Step(`^I inspect the Office link identity$`, func() error {
		switch format {
		case "xlsx":
			s, err := spreadsheet.OpenEditing(source, packaging.Limits{})
			failure = err
			if err == nil {
				_, failure = s.FindNumber("Sheet", "A1")
				found = failure == nil
			}
		case "pptx":
			s, err := presentation.OpenEditing(source, packaging.Limits{})
			failure = err
			if err == nil {
				_, failure = s.FindText("ppt/slides/slide1.xml", 1, "Text")
				found = failure == nil
			}
		}
		return nil
	})
	sc.Step(`^exactly the intended sheet or slide is found without mutation$`, func() error {
		if failure != nil || !found {
			return fmt.Errorf("correct URI selection failed: %v", failure)
		}
		if !bytes.Equal(original, source) {
			return fmt.Errorf("input changed")
		}
		return nil
	})
	sc.Step(`^the Office link is refused and input bytes are unchanged$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("incorrect identity accepted: %v", failure)
		}
		if !bytes.Equal(original, source) {
			return fmt.Errorf("input changed")
		}
		return nil
	})
}
