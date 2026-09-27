package spreadsheet

import (
	"bytes"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func calcChainFixture(t *testing.T, defect string) *EditSession {
	t.Helper()
	q := packaging.New()
	_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`))
	_, _ = q.AddPart("xl/sheet.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData></worksheet>`))
	chain := `<calcChain xmlns="` + packaging.NSSpreadsheetML + `"><c r="B1" i="1"/></calcChain>`
	if defect == "mixed" {
		chain = strings.Replace(chain, `<c r=`, `unexpected<c r=`, 1)
	}
	_, _ = q.AddPart("custom/order.xml", packaging.ContentTypeCalcChain, []byte(chain))
	q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
	q.AddRelationship("xl/workbook.xml", "sheet.xml", packaging.RelTypeWorksheet)
	target := "../custom/order.xml"
	if defect == "missing" {
		target = "../custom/missing.xml"
	}
	q.AddRelationship("xl/workbook.xml", target, packaging.RelTypeCalcChain)
	if defect == "signed" {
		_, _ = q.AddPart("_xmlsignatures/sig.xml", packaging.ContentTypeXML, []byte(`<signature/>`))
	}
	var b bytes.Buffer
	if err := q.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	s, err := OpenEditing(b.Bytes(), packaging.Limits{})
	if defect == "missing" {
		if err == nil {
			t.Fatal("unresolved chain graph accepted")
		}
		return nil
	}
	if err != nil {
		t.Fatal(err)
	}
	return s
}
func TestCalculationChainLifecycleBatch(t *testing.T) {
	t.Run("nonstandard owned chain removal then repeated edit", func(t *testing.T) {
		s := calcChainFixture(t, "")
		target, err := s.FindNumber("S", "A1")
		if err != nil {
			t.Fatal(err)
		}
		effect, err := s.SetNumberWithInvalidation(target, 3)
		if err != nil || effect.RemovedCalculationChain != "custom/order.xml" {
			t.Fatal(effect, err)
		}
		r := s.pkg.Receipt()
		found := false
		for _, c := range r.Changes {
			if c.Part == "custom/order.xml" && c.Operation == "delete" && c.AfterSHA256 == "" {
				found = true
			}
		}
		if !found {
			t.Fatal("missing deletion effect")
		}
		if _, err = s.pkg.Graph(); err != nil {
			t.Fatal(err)
		}
		target, err = s.FindNumber("S", "A1")
		if err != nil {
			t.Fatal(err)
		}
		effect, err = s.SetNumberWithInvalidation(target, 4)
		if err != nil || effect.RemovedCalculationChain != "" {
			t.Fatal(effect, err)
		}
	})
	for _, defect := range []string{"mixed", "signed"} {
		t.Run(defect, func(t *testing.T) {
			s := calcChainFixture(t, defect)
			target, _ := s.FindNumber("S", "A1")
			var before bytes.Buffer
			_ = s.pkg.WriteTo(&before)
			if _, err := s.SetNumberWithInvalidation(target, 3); err == nil {
				t.Fatal("unproved chain mutation accepted")
			}
			var after bytes.Buffer
			_ = s.pkg.WriteTo(&after)
			if !bytes.Equal(before.Bytes(), after.Bytes()) || target.consumed {
				t.Fatal("late refusal consumed source/handle")
			}
		})
	}
	t.Run("unresolved relationship rejected before policy helper", func(t *testing.T) { _ = calcChainFixture(t, "missing") })
}
