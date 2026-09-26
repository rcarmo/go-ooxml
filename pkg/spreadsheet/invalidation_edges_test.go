package spreadsheet

import (
	"bytes"
	"errors"
	"math"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestInvalidationEdgeBatch(t *testing.T) {
	setup := func(expression string) *EditSession {
		t.Helper()
		q := packaging.New()
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheets><sheet name="O'Brien" sheetId="1" r:id="rId1"/><sheet name="Calc" sheetId="2" r:id="rId2"/></sheets><calcPr calcMode="auto"/></workbook>`))
		_, _ = q.AddPart("xl/a.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`))
		_, _ = q.AddPart("xl/b.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1"><c r="A1"><f>`+expression+`</f><v>2</v></c></row></sheetData></worksheet>`))
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "a.xml", packaging.RelTypeWorksheet)
		q.AddRelationship("xl/workbook.xml", "b.xml", packaging.RelTypeWorksheet)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			t.Fatal(err)
		}
		s, err := OpenEditing(b.Bytes(), packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		return s
	}
	t.Run("quoted sheet and huge static range", func(t *testing.T) {
		s := setup(`SUM('O''Brien'!$A$1:$XFD$1048576)`)
		target, _ := s.FindNumber("O'Brien", "A1")
		effect, err := s.SetNumberWithInvalidation(target, 7)
		if err != nil || len(effect.Invalidated) != 1 {
			t.Fatal(effect, err)
		}
		errTarget := target
		_, err = s.SetNumberWithInvalidation(errTarget, 8)
		var refusal *packaging.Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "stale_target" {
			t.Fatal("old target not stale", err)
		}
	})
	t.Run("nonfinite refusal does not consume target", func(t *testing.T) {
		s := setup(`'O''Brien'!A1*2`)
		target, _ := s.FindNumber("O'Brien", "A1")
		before, _, _ := s.pkg.Part("xl/a.xml")
		for _, v := range []float64{math.NaN(), math.Inf(1), math.Inf(-1)} {
			if _, err := s.SetNumberWithInvalidation(target, v); err == nil {
				t.Fatal("nonfinite accepted")
			}
			after, _, _ := s.pkg.Part("xl/a.xml")
			if !bytes.Equal(before, after) || target.consumed {
				t.Fatal("refusal mutated")
			}
		}
		if _, err := s.SetNumberWithInvalidation(target, 4); err != nil {
			t.Fatal(err)
		}
	})
	t.Run("foreign target refuses", func(t *testing.T) {
		a, b := setup(`'O''Brien'!A1*2`), setup(`'O''Brien'!A1*2`)
		target, _ := a.FindNumber("O'Brien", "A1")
		if _, err := b.SetNumberWithInvalidation(target, 4); err == nil {
			t.Fatal("foreign target accepted")
		}
	})
}
