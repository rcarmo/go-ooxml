package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func cacheSteps(sc *godog.ScenarioContext) {
	var s *spreadsheet.EditSession
	var source, output []byte
	var result spreadsheet.CalculationEffect
	var failure error
	setup := func(condition string) error {
		q := packaging.New()
		calcPr := `<calcPr calcId="123"/>`
		protection := ""
		if condition == "inactive workbook protection" {
			protection = `<workbookProtection/><definedNames/>`
		}
		if condition == "locked workbook" {
			protection = `<workbookProtection lockStructure="1"/>`
		}
		if condition == "manual calculation mode" {
			calcPr = `<calcPr calcMode="manual"/>`
		}
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`">`+protection+`<sheets><sheet name="Input" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rId2"/></sheets>`+calcPr+`</workbook>`))
		f := "Input!A1*2"
		fAttrs := ""
		if condition == "dynamic formula" {
			f = `INDIRECT(&quot;Input!A1&quot;)*2`
		}
		if condition == "shared formula" {
			fAttrs = ` t="shared" si="0" ref="A1:A2"`
		}
		if condition == "unknown sheet" {
			f = "Missing!A1*2"
		}
		unrelated := `<c r="C1"><v>5</v></c>`
		if condition == "duplicate unrelated value" {
			unrelated = `<c r="C1"><v>5</v><v>6</v></c>`
		}
		if condition == "unclassified cell metadata" {
			unrelated = `<c r="C1" cm="1"><v>5</v></c>`
		}
		extraData := ""
		if condition == "duplicate sheetData containers" {
			extraData = `<sheetData><row r="2"><c r="D2"><v>8</v></c></row></sheetData>`
		}
		_, _ = q.AddPart("xl/worksheets/sheet1.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1"><c r="A1"><v>1</v></c>`+unrelated+`</row></sheetData>`+extraData+`</worksheet>`))
		_, _ = q.AddPart("xl/worksheets/sheet2.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1"><c r="A1"><f`+fAttrs+`>`+f+`</f><v>2</v></c><c r="B1"><f>A1+1</f><v>3</v></c><c r="C1"><f>42</f><v>42</v></c></row></sheetData></worksheet>`))
		_, _ = q.AddPart("custom/opaque.bin", "application/octet-stream", []byte("sentinel"))
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "worksheets/sheet1.xml", packaging.RelTypeWorksheet)
		q.AddRelationship("xl/workbook.xml", "worksheets/sheet2.xml", packaging.RelTypeWorksheet)
		if condition == "existing calculation chain" {
			_, _ = q.AddPart("xl/calcChain.xml", packaging.ContentTypeXML, []byte(`<calcChain xmlns="`+packaging.NSSpreadsheetML+`"/>`))
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
	save := func() error {
		dir, err := os.MkdirTemp("", "cache-")
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
	apply := func(cell string) error {
		target, err := s.FindNumber("Input", cell)
		if err != nil {
			return err
		}
		result, err = s.SetNumberWithInvalidation(target, 10)
		if err != nil {
			return err
		}
		return save()
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		source = nil
		output = nil
		failure = nil
		result = spreadsheet.CalculationEffect{}
		return ctx, nil
	})
	sc.Step(`^a workbook with input one and a static cross-sheet formula chain$`, func() error { return setup("") })
	sc.Step(`^a formula workbook containing "([^"]+)"$`, setup)
	sc.Step(`^a dependency workbook with "([^"]+)"$`, setup)
	sc.Step(`^I set the input to ten with explicit invalidation$`, func() error { return apply("A1") })
	sc.Step(`^direct and transitive cached results are absent or empty$`, func() error {
		parts, err := zipPayloads(output)
		if err != nil {
			return err
		}
		for _, address := range []string{"A1", "B1"} {
			value, _, err := sharedCell(parts["xl/worksheets/sheet2.xml"], address)
			if err != nil {
				return err
			}
			if value != "" {
				return fmt.Errorf("stale cache %s=%s", address, value)
			}
		}
		input, _, err := sharedCell(parts["xl/worksheets/sheet1.xml"], "A1")
		if err != nil {
			return err
		}
		if input != "10" {
			return fmt.Errorf("input not changed")
		}
		return nil
	})
	sc.Step(`^unrelated cached results and all formula text remain unchanged$`, func() error {
		before, _ := zipPayloads(source)
		after, err := zipPayloads(output)
		if err != nil {
			return err
		}
		want := bytes.Replace(before["xl/worksheets/sheet2.xml"], []byte(`<v>2</v>`), []byte(`<v></v>`), 1)
		want = bytes.Replace(want, []byte(`<v>3</v>`), []byte(`<v></v>`), 1)
		if !bytes.Equal(want, after["xl/worksheets/sheet2.xml"]) || !bytes.Equal(before["custom/opaque.bin"], after["custom/opaque.bin"]) {
			return fmt.Errorf("unrelated content changed")
		}
		return nil
	})
	sc.Step(`^the workbook requests full recalculation without computing an answer$`, func() error {
		if result.State != "recalculation-required" || !result.ValueChanged || len(result.Invalidated) != 2 {
			return fmt.Errorf("effect %+v", result)
		}
		parts, _ := zipPayloads(output)
		d, err := losslessxml.Parse(parts["xl/workbook.xml"])
		if err != nil {
			return err
		}
		for _, e := range d.Elements() {
			if e.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "calcPr"}) {
				attrs := map[string]string{}
				for _, a := range e.Attributes() {
					attrs[a.Name.Local] = a.Value
				}
				if attrs["calcMode"] == "auto" && attrs["fullCalcOnLoad"] == "1" && attrs["forceFullCalc"] == "1" {
					return nil
				}
			}
		}
		return fmt.Errorf("recalculation flags absent")
	})
	sc.Step(`^I change an unrelated numeric cell with explicit invalidation$`, func() error { return apply("C1") })
	sc.Step(`^cache payloads and workbook calculation metadata remain byte-identical$`, func() error {
		before, _ := zipPayloads(source)
		after, err := zipPayloads(output)
		if err != nil {
			return err
		}
		for _, p := range []string{"xl/workbook.xml", "xl/worksheets/sheet2.xml"} {
			if !bytes.Equal(before[p], after[p]) {
				return fmt.Errorf("unrelated part changed %s", p)
			}
		}
		if !result.ValueChanged || len(result.Invalidated) != 0 {
			return fmt.Errorf("unexpected effect %+v", result)
		}
		return nil
	})
	sc.Step(`^I attempt an explicit invalidating input edit$`, func() error {
		target, err := s.FindNumber("Input", "A1")
		if err != nil {
			return err
		}
		result, failure = s.SetNumberWithInvalidation(target, 10)
		return nil
	})
	sc.Step(`^invalidation returns a typed refusal without any package mutation$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("expected refusal, got %v", failure)
		}
		if err := save(); err != nil {
			return err
		}
		if !bytes.Equal(source, output) {
			return fmt.Errorf("partial edit after refusal")
		}
		return nil
	})
}
