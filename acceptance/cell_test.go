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
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func cellSteps(sc *godog.ScenarioContext) {
	var source, worksheet []byte
	var s *spreadsheet.EditSession
	var target *spreadsheet.NumberTarget
	var failure error
	var dir string
	setup := func(condition string) error {
		cells := `<c r="B2" s="0"><v>7</v></c>`
		if condition == "unknown cell attribute namespace" {
			cells = `<c xmlns:x="urn:extension" r="B2" x:formula="dependent"><v>7</v></c>`
		}
		extra := ""
		if condition == "unknown worksheet element" {
			extra = `<futureDependency ref="B2"/>`
		}
		if condition == "duplicate cell address" {
			cells += cells
		}
		if condition == "protected worksheet" {
			extra = `<sheetProtection sheet="1"/>`
		}
		if condition == "data validation" {
			extra = `<dataValidations count="1"><dataValidation sqref="B2" type="whole"><formula1>1</formula1></dataValidation></dataValidations>`
		}
		worksheet = []byte(`<worksheet xmlns="` + packaging.NSSpreadsheetML + `"><sheetData><row r="2">` + cells + `</row></sheetData>` + extra + `</worksheet>`)
		names := ""
		if condition == "defined name" {
			names = `<definedNames><definedName name="Input">Sheet1!$B$2</definedName></definedNames>`
		}
		q := packaging.New()
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSRelationships+`"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets>`+names+`</workbook>`))
		_, _ = q.AddPart("xl/worksheets/sheet1.xml", packaging.ContentTypeWorksheet, worksheet)
		second := `<sheetData/>`
		if condition == "formula on another sheet" {
			second = `<sheetData><row r="1"><c r="A1"><f>Sheet1!B2*2</f><v>14</v></c></row></sheetData>`
		}
		_, _ = q.AddPart("xl/worksheets/sheet2.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`">`+second+`</worksheet>`))
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "worksheets/sheet1.xml", packaging.RelTypeWorksheet)
		q.AddRelationship("xl/workbook.xml", "worksheets/sheet2.xml", packaging.RelTypeWorksheet)
		_, _ = q.AddPart("custom/opaque.bin", "application/octet-stream", []byte("unchanged"))
		if condition == "chart dependency" {
			_, _ = q.AddPart("xl/charts/chart1.xml", packaging.ContentTypeXML, []byte(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`))
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
	save := func() ([]byte, error) {
		var err error
		dir, err = os.MkdirTemp("", "cell-edit-")
		if err != nil {
			return nil, err
		}
		path := filepath.Join(dir, "out.xlsx")
		if _, err = s.SaveAs(path); err != nil {
			return nil, err
		}
		return os.ReadFile(path)
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		dir = ""
		s = nil
		target = nil
		failure = nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if dir != "" {
			return ctx, os.RemoveAll(dir)
		}
		return ctx, nil
	})
	sc.Step(`^a formula-free workbook with numeric cell "B2"$`, func() error { return setup("") })
	sc.Step(`^a workbook containing "([^"]+)"$`, setup)
	sc.Step(`^I set the guarded numeric cell to (\d+)$`, func(value int) error {
		var err error
		target, err = s.FindNumber("Sheet1", "B2")
		if err != nil {
			return err
		}
		return s.SetNumber(target, float64(value))
	})
	sc.Step(`^delivery changes only the cell value bytes$`, func() error {
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
			return fmt.Errorf("part set changed")
		}
		for name, b := range before {
			want := b
			if name == "xl/worksheets/sheet1.xml" {
				want = []byte(strings.Replace(string(worksheet), ">7<", ">125<", 1))
			}
			if !bytes.Equal(want, after[name]) {
				return fmt.Errorf("unexpected payload %s", name)
			}
		}
		return nil
	})
	sc.Step(`^the consumed cell target refuses reuse$`, func() error {
		err := s.SetNumber(target, 126)
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("got %v", err)
		}
		return nil
	})
	sc.Step(`^I attempt the guarded numeric correction$`, func() error {
		target, failure = s.FindNumber("Sheet1", "B2")
		if failure == nil {
			failure = s.SetNumber(target, 125)
		}
		return nil
	})
	sc.Step(`^the numeric correction returns a typed refusal$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected refusal, got %v", failure)
		}
		return nil
	})
	sc.Step(`^saving the numeric session preserves the original archive$`, func() error {
		data, err := save()
		if err != nil {
			return err
		}
		if !bytes.Equal(data, source) {
			return fmt.Errorf("refusal changed archive")
		}
		return nil
	})
}
