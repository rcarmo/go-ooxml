package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func vocabularySteps(sc *godog.ScenarioContext) {
	var s *spreadsheet.EditSession
	var source []byte
	var got, want []spreadsheet.ValidationValue
	var failure error
	str := func(v string) spreadsheet.ValidationValue {
		return spreadsheet.ValidationValue{Kind: "string", Text: v}
	}
	blank := spreadsheet.ValidationValue{Kind: "blank"}
	escape := func(v string) string { var b bytes.Buffer; _ = xml.EscapeText(&b, []byte(v)); return b.String() }
	unchanged := func() error {
		dir, err := os.MkdirTemp("", "vocab-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		p := filepath.Join(dir, "out.xlsx")
		receipt, err := s.SaveAs(p)
		if err != nil {
			return err
		}
		b, err := os.ReadFile(p)
		if err != nil {
			return err
		}
		if len(receipt.Changes) != 0 || !bytes.Equal(source, b) {
			return fmt.Errorf("inspection changed package")
		}
		return nil
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = nil
		source = nil
		got = nil
		want = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a validation workbook with "([^"]+)"$`, func(condition string) error {
		expr := `" Yes,No ,A""B,"`
		sqref := "D1"
		kind := "list"
		extraValidation, merge, extension := "", "", ""
		cells := `<row r="1"><c r="A1" t="inlineStr"><is><t>Yes</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t xml:space="preserve"> No </t></is></c></row>`
		want = []spreadsheet.ValidationValue{str(" Yes"), str("No "), str(`A"B`), str("")}
		shared := ""
		sharedRels := 1
		switch condition {
		case "literal whitespace quotes and empty item":
		case "no matching validation":
			sqref = "E1"
			want = nil
		case "non-list validation":
			kind = "whole"
			want = nil
		case "reversed local range with blank":
			expr = "=$A$3:$A$1"
			want = []spreadsheet.ValidationValue{str("Yes"), blank, str(" No ")}
		case "quoted cross-sheet range":
			expr = "='Owner''s Inputs'!$B$2:$B$1"
			want = []spreadsheet.ValidationValue{str("Low"), str("High")}
		case "overlapping sqref areas in one list":
			sqref = "C1:E1 D1:D2"
		case "boolean numeric and empty string":
			expr = "A1:A3"
			cells = `<row r="1"><c r="A1" t="b"><v>1</v></c></row><row r="2"><c r="A2"><v>12.50</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t/></is></c></row>`
			want = []spreadsheet.ValidationValue{{Kind: "boolean", Text: "1"}, {Kind: "number", Text: "12.50"}, str("")}
		case "shared strings with duplicate entries", "out of range shared index", "missing shared string relationship", "duplicate shared string relationships", "mixed shared string structure", "rich shared string":
			expr = "A1:A3"
			cells = `<row r="1"><c r="A1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c></row><row r="3"><c r="A3" t="s"><v>0</v></c></row>`
			shared = `<sst xmlns="` + packaging.NSSpreadsheetML + `" count="3" uniqueCount="3"><si><t>Dup</t></si><si><t>Dup</t></si><si><t>Last</t></si></sst>`
			want = []spreadsheet.ValidationValue{str("Dup"), str("Last"), str("Dup")}
			if condition == "out of range shared index" {
				cells = strings.Replace(cells, `<v>1</v>`, `<v>9</v>`, 1)
			}
			if condition == "missing shared string relationship" {
				sharedRels = 0
			}
			if condition == "duplicate shared string relationships" {
				sharedRels = 2
			}
			if condition == "mixed shared string structure" {
				shared = strings.Replace(shared, `<si><t>Last</t></si>`, `<si><t>Last</t><r><t>extra</t></r></si>`, 1)
			}
			if condition == "rich shared string" {
				shared = strings.Replace(shared, `<si><t>Last</t></si>`, `<si><r><rPr><b/></rPr><t xml:space="preserve"> rich </t></r><r><t>text</t></r></si>`, 1)
				want[1] = str(" rich text")
			}
		case "rich inline string", "missing rich run text":
			expr = "A1"
			cells = `<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t xml:space="preserve"> rich </t></r><r><t>text</t></r></is></c></row>`
			want = []spreadsheet.ValidationValue{str(" rich text")}
			if condition == "missing rich run text" {
				cells = strings.Replace(cells, `<r><t>text</t></r>`, `<r><rPr><i/></rPr></r>`, 1)
			}
		case "two matching lists":
			extraValidation = `<dataValidation type="list" sqref="D1"><formula1>&quot;X,Y&quot;</formula1></dataValidation>`
		case "invalid literal quote":
			expr = `"A"B,C"`
		case "dynamic source":
			expr = `INDIRECT("A1:A3")`
		case "absent source sheet":
			expr = "Missing!A1:A3"
		case "two-dimensional source":
			expr = "A1:B3"
		case "source formula with cached value":
			expr = "A1:A3"
			cells = `<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>`
		case "source error":
			expr = "A1"
			cells = `<row r="1"><c r="A1" t="e"><v>#REF!</v></c></row>`
		case "merged source interior":
			expr = "A1:A3"
			merge = `<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>`
		case "merged target interior":
			merge = `<mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>`
		case "extended validation":
			extension = `<extLst><ext uri="unknown"><x:dataValidations xmlns:x="urn:extension"/></ext></extLst>`
		case "duplicate source cell":
			expr = "A1"
			cells = `<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row>`
		case "arithmetic after range":
			expr = "A1:A3+1"
		default:
			return fmt.Errorf("unknown condition %s", condition)
		}
		count := "1"
		if extraValidation != "" {
			count = "2"
		}
		validation := `<dataValidations count="` + count + `"><dataValidation type="` + kind + `" allowBlank="1" sqref="` + sqref + `"><formula1>` + escape(expr) + `</formula1></dataValidation>` + extraValidation + `</dataValidations>`
		q := packaging.New()
		ns := packaging.NSSpreadsheetML
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+ns+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheets><sheet name="Main" sheetId="1" r:id="rId1"/><sheet name="Owner's Inputs" sheetId="2" r:id="rId2"/></sheets></workbook>`))
		_, _ = q.AddPart("xl/sheet1.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+ns+`"><sheetData>`+cells+`</sheetData>`+merge+validation+extension+`</worksheet>`))
		_, _ = q.AddPart("xl/sheet2.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+ns+`"><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Low</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>High</t></is></c></row></sheetData></worksheet>`))
		_, _ = q.AddPart("opaque.bin", "application/octet-stream", []byte("sentinel"))
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "sheet1.xml", packaging.RelTypeWorksheet)
		q.AddRelationship("xl/workbook.xml", "sheet2.xml", packaging.RelTypeWorksheet)
		if shared != "" {
			_, _ = q.AddPart("xl/sharedStrings.xml", packaging.ContentTypeSharedStrings, []byte(shared))
			for i := 0; i < sharedRels; i++ {
				q.AddRelationship("xl/workbook.xml", "sharedStrings.xml", packaging.RelTypeSharedStrings)
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
	})
	sc.Step(`^I inspect its allowed values$`, func() error { got, failure = s.AllowedValues("Main", "D1"); return nil })
	sc.Step(`^the expected validation vocabulary is returned without mutation$`, func() error {
		if failure != nil {
			return failure
		}
		if !reflect.DeepEqual(got, want) {
			return fmt.Errorf("got%+v want%+v", got, want)
		}
		return unchanged()
	})
	sc.Step(`^validation inspection refuses and returns no partial values$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) || got != nil {
			return fmt.Errorf("partial or successful vocabulary %+v %v", got, failure)
		}
		return unchanged()
	})
}
