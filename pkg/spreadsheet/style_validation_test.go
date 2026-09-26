package spreadsheet

import (
	"bytes"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestStyleReferenceBatch(t *testing.T) {
	for _, tc := range []struct {
		name, table, cells string
		bad                bool
	}{{"implicit no table", "", `<c r="A1"><v>1</v></c>`, false}, {"nonzero no table", "", `<c r="A1" s="1"><v>1</v></c>`, true}, {"empty explicit zero", `<cellXfs count="0"/>`, `<c r="A1" s="0"><v>1</v></c>`, true}, {"empty implicit zero", `<cellXfs count="0"/>`, `<c r="A1"><v>1</v></c>`, true}, {"count absent but actual xf", `<cellXfs><xf/></cellXfs>`, `<c r="A1"><v>1</v></c>`, false}, {"overflow", `<cellXfs><xf/></cellXfs>`, `<c r="A1" s="4294967296"><v>1</v></c>`, true}, {"empty index", `<cellXfs><xf/></cellXfs>`, `<c r="A1" s=""><v>1</v></c>`, true}, {"two tables", `<cellXfs><xf/></cellXfs><cellXfs><xf/></cellXfs>`, `<c r="A1"><v>1</v></c>`, true}} {
		t.Run(tc.name, func(t *testing.T) {
			q := packaging.New()
			_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`))
			_, _ = q.AddPart("xl/worksheets/sheet1.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`"><sheetData><row r="1">`+tc.cells+`</row></sheetData></worksheet>`))
			q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
			q.AddRelationship("xl/workbook.xml", "worksheets/sheet1.xml", packaging.RelTypeWorksheet)
			if tc.table != "" {
				_, _ = q.AddPart("xl/styles.xml", packaging.ContentTypeExcelStyles, []byte(`<styleSheet xmlns="`+packaging.NSSpreadsheetML+`">`+tc.table+`</styleSheet>`))
				q.AddRelationship("xl/workbook.xml", "styles.xml", packaging.RelTypeStyles)
			}
			var b bytes.Buffer
			if err := q.WriteTo(&b); err != nil {
				t.Fatal(err)
			}
			s, err := OpenEditing(b.Bytes(), packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if err = s.ValidateStyles(); (err != nil) != tc.bad {
				t.Fatalf("bad=%v error=%v", tc.bad, err)
			}
		})
	}
}
