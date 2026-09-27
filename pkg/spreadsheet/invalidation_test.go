package spreadsheet

import (
	"bytes"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestStaticInvalidationBatch(t *testing.T) {
	open := func(formulas string) *EditSession {
		t.Helper()
		q := packaging.New()
		_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`))
		_, _ = q.AddPart("xl/sheet.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1"><c r="A1"><v>1</v></c>`+formulas+`</row></sheetData></worksheet>`))
		q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationship("xl/workbook.xml", "sheet.xml", packaging.RelTypeWorksheet)
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
	t.Run("range same sheet empty cache and no-op", func(t *testing.T) {
		s := open(`<c r="B1"><f>SUM($A$1:A3)</f><v>1</v></c><c r="C1"><f>B1+1</f><v/></c>`)
		target, _ := s.FindNumber("S", "A1")
		before, _, _ := s.pkg.Part("xl/sheet.xml")
		effect, err := s.SetNumberWithInvalidation(target, 1)
		if err != nil || effect.ValueChanged {
			t.Fatal(effect, err)
		}
		after, _, _ := s.pkg.Part("xl/sheet.xml")
		if !bytes.Equal(before, after) {
			t.Fatal("no-op mutated")
		}
		effect, err = s.SetNumberWithInvalidation(target, 5)
		if err != nil || len(effect.Invalidated) != 2 {
			t.Fatal(effect, err)
		}
		main, _, _ := s.pkg.Part(s.main)
		if !bytes.Contains(main, []byte(`fullCalcOnLoad="1"`)) {
			t.Fatal("missing appended calculation flags")
		}
	})
	t.Run("cycles terminate with affected closure", func(t *testing.T) {
		s := open(`<c r="B1"><f>A1+C1</f><v>2</v></c><c r="C1"><f>B1</f><v>2</v></c>`)
		target, _ := s.FindNumber("S", "A1")
		effect, err := s.SetNumberWithInvalidation(target, 3)
		if err != nil || len(effect.Invalidated) != 2 {
			t.Fatal(effect, err)
		}
	})
	t.Run("string literals not references", func(t *testing.T) {
		s := open(`<c r="B1" t="str"><f>&quot;A1&quot;</f><v>A1</v></c>`)
		target, _ := s.FindNumber("S", "A1")
		effect, err := s.SetNumberWithInvalidation(target, 3)
		if err != nil || len(effect.Invalidated) != 0 || effect.State != "caches-unchanged" {
			t.Fatal(effect, err)
		}
	})
	t.Run("array and volatile refusal reusable", func(t *testing.T) {
		for _, formulaXML := range []string{`<f t="array" ref="B1:B2">A1*2</f>`, `<f>NOW()+A1</f>`} {
			s := open(`<c r="B1">` + formulaXML + `<v>2</v></c>`)
			target, _ := s.FindNumber("S", "A1")
			before, _, _ := s.pkg.Part("xl/sheet.xml")
			if _, err := s.SetNumberWithInvalidation(target, 3); err == nil {
				t.Fatal("unsupported accepted")
			}
			after, _, _ := s.pkg.Part("xl/sheet.xml")
			if !bytes.Equal(before, after) || target.consumed {
				t.Fatal("refusal mutated state")
			}
		}
	})
}
