package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func cacheSteps(sc *godog.ScenarioContext) {
	var s *spreadsheet.EditSession
	var source, output, recordedSource []byte
	var canonicalPath string
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
		if condition != "existing calculation chain" && strings.Contains(condition, "calculation chain") || condition == "chain with outgoing relationships" {
			chain := `<calcChain xmlns="` + packaging.NSSpreadsheetML + `"><c r="A1" i="2"/><c r="B1"/></calcChain>`
			if condition == "extended calculation chain" {
				chain = `<calcChain xmlns="` + packaging.NSSpreadsheetML + `"><extLst/></calcChain>`
			}
			if condition == "external calculation chain" {
				q.AddRelationshipWithTargetMode("xl/workbook.xml", "https://example.invalid/chain.xml", packaging.RelTypeCalcChain, packaging.TargetModeExternal)
			} else {
				_, _ = q.AddPart("xl/chains/order.xml", packaging.ContentTypeCalcChain, []byte(chain))
				q.AddRelationship("xl/workbook.xml", "chains/order.xml", packaging.RelTypeCalcChain)
				if condition == "shared calculation chain" {
					q.AddRelationship("xl/worksheets/sheet1.xml", "../chains/order.xml", packaging.RelTypeCalcChain)
				}
				if condition == "chain with outgoing relationships" {
					q.AddRelationship("xl/chains/order.xml", "../../custom/opaque.bin", packaging.RelTypeImage)
				}
			}
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
		recordedSource = nil
		canonicalPath = ""
		failure = nil
		result = spreadsheet.CalculationEffect{}
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if canonicalPath != "" {
			_ = os.RemoveAll(filepath.Dir(canonicalPath))
		}
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
	sc.Step(`^the calculation chain part and its registrations are removed$`, func() error {
		parts, err := zipPayloads(output)
		if err != nil {
			return err
		}
		if _, ok := parts["xl/chains/order.xml"]; ok {
			return fmt.Errorf("chain retained")
		}
		q, err := packaging.OpenPreserved(output, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		for _, e := range g.Edges {
			if e.Type == packaging.RelTypeCalcChain {
				return fmt.Errorf("chain edge retained")
			}
		}
		if bytes.Contains(parts["[Content_Types].xml"], []byte("/xl/chains/order.xml")) {
			return fmt.Errorf("chain override retained")
		}
		return nil
	})
	sc.Step(`^the calculation chain and registry payloads remain byte-identical$`, func() error {
		before, _ := zipPayloads(source)
		after, err := zipPayloads(output)
		if err != nil {
			return err
		}
		for _, name := range []string{"xl/chains/order.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"} {
			if !bytes.Equal(before[name], after[name]) {
				return fmt.Errorf("unrelated chain part changed %s", name)
			}
		}
		return nil
	})
	sc.Step(`^I apply a same-value edit then reuse its numeric target$`, func() error {
		target, err := s.FindNumber("Input", "A1")
		if err != nil {
			return err
		}
		result, err = s.SetNumberWithInvalidation(target, 1)
		if err != nil {
			return err
		}
		if err = save(); err != nil {
			return err
		}
		if !bytes.Equal(source, output) {
			return fmt.Errorf("no-op changed chain")
		}
		result, err = s.SetNumberWithInvalidation(target, 10)
		if err != nil {
			return err
		}
		return save()
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

	// The selected canonical outcome uses the same synthetic producer as CHAIN-001,
	// but asserts custody and an independent disk readback before receiving credit.
	readSaved := func() (map[string][]byte, error) {
		if canonicalPath == "" {
			return nil, fmt.Errorf("canonical destination not saved")
		}
		b, err := os.ReadFile(canonicalPath)
		if err != nil {
			return nil, err
		}
		return zipPayloads(b)
	}
	cell := func(part, address, wantValue, wantFormula string) error {
		parts, err := readSaved()
		if err != nil {
			return err
		}
		value, formula, err := sharedCell(parts[part], address)
		if err != nil {
			return err
		}
		if value != wantValue || formula != wantFormula {
			return fmt.Errorf("%s:%s value/formula %q/%q, want %q/%q", part, address, value, formula, wantValue, wantFormula)
		}
		return nil
	}
	sc.Step(`^a two-sheet XLSX package with Input!A1 numeric 1 and Calc!A1 formula "Input!A1\*2" cached as 2$`, func() error {
		if err := setup("owned calculation chain"); err != nil {
			return err
		}
		parts, err := zipPayloads(source)
		if err != nil {
			return err
		}
		for _, c := range []struct{ part, address, value, formula string }{
			{"xl/worksheets/sheet1.xml", "A1", "1", ""},
			{"xl/worksheets/sheet2.xml", "A1", "2", "Input!A1*2"},
		} {
			v, f, err := sharedCell(parts[c.part], c.address)
			if err != nil || v != c.value || f != c.formula {
				return fmt.Errorf("initial %s:%s value/formula %q/%q: %v", c.part, c.address, v, f, err)
			}
		}
		return nil
	})
	sc.Step(`^Calc!B1 formula "A1\+1" is cached as 3 and unrelated Calc!C1 formula "42" is cached as 42$`, func() error {
		parts, err := zipPayloads(source)
		if err != nil {
			return err
		}
		for _, c := range []struct{ address, value, formula string }{{"B1", "3", "A1+1"}, {"C1", "42", "42"}} {
			v, f, err := sharedCell(parts["xl/worksheets/sheet2.xml"], c.address)
			if err != nil || v != c.value || f != c.formula {
				return fmt.Errorf("initial Calc!%s value/formula %q/%q: %v", c.address, v, f, err)
			}
		}
		return nil
	})
	sc.Step(`^the workbook alone owns a calculation-chain relationship to xl/chains/order.xml$`, func() error {
		q, err := packaging.OpenPreserved(source, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		count := 0
		for _, e := range g.Edges {
			if e.Type == packaging.RelTypeCalcChain {
				if e.Source != "xl/workbook.xml" || e.External || e.ResolvedPart != "xl/chains/order.xml" {
					return fmt.Errorf("unowned calculation chain: %+v", e)
				}
				count++
			}
		}
		if count != 1 {
			return fmt.Errorf("expected one workbook-owned chain edge, got %d", count)
		}
		return nil
	})
	sc.Step(`^xl/chains/order.xml has the calculation-chain content type and entries for Calc!A1 and Calc!B1$`, func() error {
		q, err := packaging.OpenPreserved(source, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		count := 0
		for _, part := range g.Parts {
			if part.Name == "xl/chains/order.xml" {
				if part.ContentType != packaging.ContentTypeCalcChain {
					return fmt.Errorf("wrong chain content type %s", part.ContentType)
				}
				count++
			}
		}
		if count != 1 {
			return fmt.Errorf("expected one chain part, got %d", count)
		}
		parts, err := zipPayloads(source)
		if err != nil {
			return err
		}
		d, err := losslessxml.Parse(parts["xl/chains/order.xml"])
		if err != nil {
			return err
		}
		var addresses []string
		for _, e := range d.Elements() {
			if e.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "c"}) {
				continue
			}
			for _, a := range e.Attributes() {
				if a.Name.Local == "r" {
					addresses = append(addresses, a.Value)
				}
			}
		}
		if len(addresses) != 2 || addresses[0] != "A1" || addresses[1] != "B1" {
			return fmt.Errorf("unexpected chain entries %v", addresses)
		}
		return nil
	})
	sc.Step(`^all source package member payloads and bytes are recorded$`, func() error {
		if _, err := zipPayloads(source); err != nil {
			return err
		}
		recordedSource = bytes.Clone(source)
		return nil
	})
	sc.Step(`^Input!A1 is set to numeric 10 with explicit dependent-cache invalidation and the result is saved and reopened$`, func() error {
		if recordedSource == nil {
			return fmt.Errorf("source not recorded")
		}
		target, err := s.FindNumber("Input", "A1")
		if err != nil {
			return err
		}
		result, err = s.SetNumberWithInvalidation(target, 10)
		if err != nil {
			return err
		}
		dir, err := os.MkdirTemp("", "owned-chain-")
		if err != nil {
			return err
		}
		canonicalPath = filepath.Join(dir, "out.xlsx")
		if _, err := s.SaveAs(canonicalPath); err != nil {
			return err
		}
		saved, err := os.ReadFile(canonicalPath)
		if err != nil {
			return err
		}
		reopened, err := spreadsheet.OpenEditing(saved, packaging.Limits{})
		if err != nil {
			return err
		}
		_, err = reopened.FindNumber("Input", "A1")
		return err
	})
	sc.Step(`^the source package bytes remain unchanged and reopened Input!A1 is numeric 10$`, func() error {
		if !bytes.Equal(source, recordedSource) {
			return fmt.Errorf("source bytes changed")
		}
		return cell("xl/worksheets/sheet1.xml", "A1", "10", "")
	})
	sc.Step(`^the reopened Calc!A1 and Calc!B1 formulas are unchanged with absent or empty cached values$`, func() error {
		if err := cell("xl/worksheets/sheet2.xml", "A1", "", "Input!A1*2"); err != nil {
			return err
		}
		return cell("xl/worksheets/sheet2.xml", "B1", "", "A1+1")
	})
	sc.Step(`^a data-only read of those cells cannot return the old cached values 2 and 3 as current$`, func() error {
		// Read the saved ZIP independently and decode only stored value leaves;
		// Go has no separate data-only spreadsheet API or formula evaluator.
		z, err := zip.OpenReader(canonicalPath)
		if err != nil {
			return err
		}
		defer z.Close()
		found := 0
		for _, f := range z.File {
			if f.Name != "xl/worksheets/sheet2.xml" {
				continue
			}
			found++
			r, err := f.Open()
			if err != nil {
				return err
			}
			b, readErr := io.ReadAll(r)
			closeErr := r.Close()
			if readErr != nil {
				return readErr
			}
			if closeErr != nil {
				return closeErr
			}
			var worksheet struct {
				Cells []struct {
					Address string `xml:"r,attr"`
					Value   string `xml:"v"`
				} `xml:"sheetData>row>c"`
			}
			if err := xml.Unmarshal(b, &worksheet); err != nil {
				return err
			}
			counts := map[string]int{}
			for _, c := range worksheet.Cells {
				if c.Address == "A1" || c.Address == "B1" {
					counts[c.Address]++
					if c.Value != "" {
						return fmt.Errorf("data-only Calc!%s exposes cached %q", c.Address, c.Value)
					}
				}
			}
			if counts["A1"] != 1 || counts["B1"] != 1 {
				return fmt.Errorf("data-only cell counts %v", counts)
			}
		}
		if found != 1 {
			return fmt.Errorf("expected one reopened Calc worksheet, got %d", found)
		}
		return nil
	})
	sc.Step(`^the reopened Calc!C1 formula and cached value 42 remain unchanged$`, func() error {
		return cell("xl/worksheets/sheet2.xml", "C1", "42", "42")
	})
	sc.Step(`^the workbook requests full recalculation without claiming a computed result$`, func() error {
		if result.State != "recalculation-required" || !result.ValueChanged || len(result.Invalidated) != 2 {
			return fmt.Errorf("calculation effect %+v", result)
		}
		parts, err := readSaved()
		if err != nil {
			return err
		}
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
	sc.Step(`^xl/chains/order.xml, its workbook relationship and its content-type override are absent$`, func() error {
		parts, err := readSaved()
		if err != nil {
			return err
		}
		if _, ok := parts["xl/chains/order.xml"]; ok {
			return fmt.Errorf("chain part retained")
		}
		b, err := os.ReadFile(canonicalPath)
		if err != nil {
			return err
		}
		q, err := packaging.OpenPreserved(b, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		for _, edge := range g.Edges {
			if edge.Type == packaging.RelTypeCalcChain {
				return fmt.Errorf("chain relationship retained")
			}
		}
		if bytes.Contains(parts["[Content_Types].xml"], []byte("/xl/chains/order.xml")) {
			return fmt.Errorf("chain content-type override retained")
		}
		return nil
	})
	sc.Step(`^every destination relationship and content-type target resolves$`, func() error {
		b, err := os.ReadFile(canonicalPath)
		if err != nil {
			return err
		}
		q, err := packaging.OpenPreserved(b, packaging.Limits{})
		if err != nil {
			return err
		}
		_, err = q.Graph() // Graph verifies local edges and content-type overrides.
		return err
	})
	sc.Step(`^every source member payload outside the workbook, two worksheets, workbook relationships and content types remains byte-identical$`, func() error {
		before, err := zipPayloads(recordedSource)
		if err != nil {
			return err
		}
		after, err := readSaved()
		if err != nil {
			return err
		}
		allowed := map[string]bool{"xl/chains/order.xml": true, "xl/workbook.xml": true, "xl/worksheets/sheet1.xml": true, "xl/worksheets/sheet2.xml": true, "xl/_rels/workbook.xml.rels": true, "[Content_Types].xml": true}
		for name, original := range before {
			if _, exists := after[name]; !exists && name != "xl/chains/order.xml" {
				return fmt.Errorf("unexpected removed member %s", name)
			}
			if !allowed[name] && !bytes.Equal(original, after[name]) {
				return fmt.Errorf("unrelated member changed %s", name)
			}
		}
		for name := range after {
			if _, exists := before[name]; !exists {
				return fmt.Errorf("unexpected added member %s", name)
			}
		}
		return nil
	})
}
