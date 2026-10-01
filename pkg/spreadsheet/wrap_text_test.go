package spreadsheet

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"testing"
)

func TestSetCellWrapTextPreservesCurrentXF(t *testing.T) {
	styles := newStyles(nil)
	styles.stylesheet.CellXfs = &sml.CellXfs{Count: 2, Xf: []*sml.Xf{{XFID: 0}, {XFID: 0, NumFmtID: 170, FontID: 2, FillID: 1, BorderID: 3, Alignment: &sml.Alignment{Horizontal: "center"}}}}
	wb := &workbookImpl{styles: styles}
	ws := &worksheetImpl{workbook: wb}
	cell := &cellImpl{worksheet: ws, cell: &sml.Cell{S: 1}}
	if err := SetCellWrapText(cell, true); err != nil {
		t.Fatal(err)
	}
	if cell.cell.S != 2 || len(styles.stylesheet.CellXfs.Xf) != 3 {
		t.Fatalf("style index/count: %d/%d", cell.cell.S, len(styles.stylesheet.CellXfs.Xf))
	}
	old := styles.stylesheet.CellXfs.Xf[1]
	current := styles.stylesheet.CellXfs.Xf[2]
	if old.Alignment.WrapText != nil || current.FontID != 2 || current.FillID != 1 || current.BorderID != 3 || current.NumFmtID != 170 || current.Alignment.Horizontal != "center" || current.Alignment.WrapText == nil || !*current.Alignment.WrapText {
		t.Fatalf("original/new style: %+v / %+v", old, current)
	}
}
