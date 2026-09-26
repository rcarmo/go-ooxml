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
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func styleIndexSteps(sc *godog.ScenarioContext) {
	var s *spreadsheet.EditSession
	var source, style, sheet []byte
	var failure error
	setup := func(condition string) error {
		attrs, extra := "", ""
		count, entries := "1", `<xf/>`
		switch condition {
		case "empty cellXfs and omitted s":
			count, entries = "0", ""
		case "empty cellXfs and explicit zero":
			count, entries, attrs = "0", "", ` s="0"`
		case "out of range style":
			attrs = ` s="1"`
		case "mismatched cellXfs count":
			count = "2"
		case "malformed style index":
			attrs = ` s="-1"`
		case "unrelated cell with invalid style":
			extra = `<c r="C2" s="9"><v>8</v></c>`
		case "one style and omitted s":
		default:
			return fmt.Errorf("unknown condition")
		}
		sheet = []byte(`<worksheet xmlns="` + packaging.NSSpreadsheetML + `"><sheetData><row r="2"><c r="B2"` + attrs + `><v>7</v></c>` + extra + `</row></sheetData></worksheet>`)
		style = []byte(`<styleSheet xmlns="` + packaging.NSSpreadsheetML + `"><cellXfs count="` + count + `">` + entries + `</cellXfs></styleSheet>`)
		q := packaging.New()
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSRelationships+`"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`))
		_, _ = q.AddPart("xl/worksheets/sheet1.xml", packaging.ContentTypeWorksheet, sheet)
		_, _ = q.AddPart("xl/styles.xml", packaging.ContentTypeExcelStyles, style)
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "worksheets/sheet1.xml", packaging.RelTypeWorksheet)
		q.AddRelationship("xl/workbook.xml", "styles.xml", packaging.RelTypeStyles)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		s, err = spreadsheet.OpenEditing(source, packaging.Limits{})
		return err
	}
	save := func() ([]byte, error) {
		dir, err := os.MkdirTemp("", "style-index-")
		if err != nil {
			return nil, err
		}
		defer os.RemoveAll(dir)
		out := filepath.Join(dir, "out.xlsx")
		if _, err = s.SaveAs(out); err != nil {
			return nil, err
		}
		return os.ReadFile(out)
	}
	apply := func() error {
		target, err := s.FindNumber("Sheet1", "B2")
		if err != nil {
			return err
		}
		return s.SetNumber(target, 125)
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a numeric workbook with "([^"]+)"$`, setup)
	sc.Step(`^I attempt a numeric edit under the style policy$`, func() error { failure = apply(); return nil })
	sc.Step(`^style validation refuses with the original archive unchanged$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected typed style refusal, got %v", failure)
		}
		got, err := save()
		if err != nil {
			return err
		}
		if !bytes.Equal(source, got) {
			return fmt.Errorf("refusal mutated archive")
		}
		return nil
	})
	sc.Step(`^I set its style-validated numeric value to 125$`, apply)
	sc.Step(`^the value changes without adding a style attribute or modifying styles$`, func() error {
		got, err := save()
		if err != nil {
			return err
		}
		parts, err := zipPayloads(got)
		if err != nil {
			return err
		}
		want := bytes.Replace(sheet, []byte(">7<"), []byte(">125<"), 1)
		if !bytes.Equal(parts["xl/worksheets/sheet1.xml"], want) || !bytes.Equal(parts["xl/styles.xml"], style) {
			return fmt.Errorf("styles or unrelated XML changed")
		}
		return nil
	})
}
