package presentation

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Table represents a table in a slide.
type tableImpl struct {
	tbl *dml.Tbl
}

func newTable(rows, cols int, width, height int64) *tableImpl {
	if rows < 1 {
		rows = 1
	}
	if cols < 1 {
		cols = 1
	}
	colWidth := width / int64(cols)
	rowHeight := height / int64(rows)

	grid := &dml.TblGrid{}
	for i := 0; i < cols; i++ {
		grid.GridCol = append(grid.GridCol, &dml.GridCol{W: colWidth})
	}

	tbl := &dml.Tbl{
		TblPr: &dml.TblPr{
			BandRow: utils.BoolPtr(true),
		},
		TblGrid: grid,
	}

	for r := 0; r < rows; r++ {
		tr := &dml.Tr{H: rowHeight}
		for c := 0; c < cols; c++ {
			tr.Tc = append(tr.Tc, &dml.Tc{
				TxBody: &dml.TxBody{
					BodyPr:   &dml.BodyPr{},
					LstStyle: &dml.LstStyle{},
					P:        []*dml.P{{R: []*dml.R{{T: ""}}}},
				},
			})
		}
		tbl.Tr = append(tbl.Tr, tr)
	}

	return &tableImpl{tbl: tbl}
}

// Rows returns table rows.
func (t *tableImpl) Rows() []TableRow {
	rows := make([]TableRow, len(t.tbl.Tr))
	for i, tr := range t.tbl.Tr {
		rows[i] = &tableRowImpl{row: tr}
	}
	return rows
}

// Row returns the row at index (0-based).
func (t *tableImpl) Row(index int) TableRow {
	if index < 0 || index >= len(t.tbl.Tr) {
		return nil
	}
	return &tableRowImpl{row: t.tbl.Tr[index]}
}

// Cell returns the cell at row/col (0-based).
func (t *tableImpl) Cell(row, col int) TableCell {
	r := t.Row(row)
	if r == nil {
		return nil
	}
	return r.Cell(col)
}

// AddRow adds a row at the end.
func (t *tableImpl) AddRow() TableRow {
	colCount := t.ColumnCount()
	tr := &dml.Tr{H: t.rowHeight()}
	for i := 0; i < colCount; i++ {
		tr.Tc = append(tr.Tc, &dml.Tc{
			TxBody: &dml.TxBody{
				BodyPr:   &dml.BodyPr{},
				LstStyle: &dml.LstStyle{},
				P:        []*dml.P{{R: []*dml.R{{T: ""}}}},
			},
		})
	}
	t.tbl.Tr = append(t.tbl.Tr, tr)
	return &tableRowImpl{row: tr}
}

// InsertRow inserts a row at index.
func (t *tableImpl) InsertRow(index int) TableRow {
	colCount := t.ColumnCount()
	tr := &dml.Tr{H: t.rowHeight()}
	for i := 0; i < colCount; i++ {
		tr.Tc = append(tr.Tc, &dml.Tc{
			TxBody: &dml.TxBody{
				BodyPr:   &dml.BodyPr{},
				LstStyle: &dml.LstStyle{},
				P:        []*dml.P{{R: []*dml.R{{T: ""}}}},
			},
		})
	}
	if index >= len(t.tbl.Tr) {
		t.tbl.Tr = append(t.tbl.Tr, tr)
	} else {
		t.tbl.Tr = append(t.tbl.Tr[:index+1], t.tbl.Tr[index:]...)
		t.tbl.Tr[index] = tr
	}
	return &tableRowImpl{row: tr}
}

// DeleteRow deletes a row at index.
func (t *tableImpl) DeleteRow(index int) error {
	if index < 0 || index >= len(t.tbl.Tr) {
		return ErrInvalidIndex
	}
	t.tbl.Tr = append(t.tbl.Tr[:index], t.tbl.Tr[index+1:]...)
	return nil
}

// RowCount returns number of rows.
func (t *tableImpl) RowCount() int {
	return len(t.tbl.Tr)
}

// ColumnCount returns number of columns.
func (t *tableImpl) ColumnCount() int {
	if t.tbl.TblGrid != nil && len(t.tbl.TblGrid.GridCol) > 0 {
		return len(t.tbl.TblGrid.GridCol)
	}
	if len(t.tbl.Tr) > 0 {
		return len(t.tbl.Tr[0].Tc)
	}
	return 0
}

// TableFromShape returns the table from a shape if present.
func TableFromShape(shape Shape) Table {
	if shape == nil || !shape.HasTable() {
		return nil
	}
	if table, ok := shape.Table().(*tableImpl); ok {
		return table
	}
	return nil
}

// XML returns the underlying DrawingML table.
func (t *tableImpl) XML() *dml.Tbl {
	return t.tbl
}

func (t *tableImpl) rowHeight() int64 {
	if len(t.tbl.Tr) > 0 {
		return t.tbl.Tr[0].H
	}
	return 0
}

// TableRow represents a table row.
type tableRowImpl struct {
	row *dml.Tr
}

// Cells returns row cells.
func (r *tableRowImpl) Cells() []TableCell {
	cells := make([]TableCell, len(r.row.Tc))
	for i, tc := range r.row.Tc {
		cells[i] = &tableCellImpl{cell: tc}
	}
	return cells
}

// Cell returns the cell at index (0-based).
func (r *tableRowImpl) Cell(index int) TableCell {
	if index < 0 || index >= len(r.row.Tc) {
		return nil
	}
	return &tableCellImpl{cell: r.row.Tc[index]}
}

// Height returns row height.
func (r *tableRowImpl) Height() int64 {
	return r.row.H
}

// SetHeight sets row height.
func (r *tableRowImpl) SetHeight(height int64) {
	r.row.H = height
}

// TableCell represents a table cell.
type tableCellImpl struct {
	cell *dml.Tc
}

// TextFrame returns the cell text frame.
func (c *tableCellImpl) TextFrame() TextFrame {
	if c.cell == nil {
		return nil
	}
	if c.cell.TxBody == nil {
		c.cell.TxBody = &dml.TxBody{
			BodyPr:   &dml.BodyPr{},
			LstStyle: &dml.LstStyle{},
			P:        []*dml.P{{R: []*dml.R{{T: ""}}}},
		}
	}
	return &textFrameImpl{txBody: c.cell.TxBody}
}

// Text returns cell text.
func (c *tableCellImpl) Text() string {
	tf := c.TextFrame()
	if tf == nil {
		return ""
	}
	if t, ok := tf.(*textFrameImpl); ok {
		return t.Text()
	}
	return ""
}

// SetText sets cell text.
func (c *tableCellImpl) SetText(text string) {
	tf := c.TextFrame()
	if tf == nil {
		return
	}
	if t, ok := tf.(*textFrameImpl); ok {
		t.SetText(text)
	}
}

// RowSpan returns row span.
func (c *tableCellImpl) RowSpan() int {
	if c.cell == nil {
		return 1
	}
	if c.cell.TcPr == nil || c.cell.TcPr.RowSpan == nil {
		return 1
	}
	return *c.cell.TcPr.RowSpan
}

// ColSpan returns column span.
func (c *tableCellImpl) ColSpan() int {
	if c.cell == nil {
		return 1
	}
	if c.cell.TcPr == nil || c.cell.TcPr.GridSpan == nil {
		return 1
	}
	return *c.cell.TcPr.GridSpan
}

// SetRowSpan sets row span.
func (c *tableCellImpl) SetRowSpan(span int) {
	if c.cell == nil {
		return
	}
	if span < 1 {
		span = 1
	}
	if c.cell.TcPr == nil {
		c.cell.TcPr = &dml.TcPr{}
	}
	c.cell.TcPr.RowSpan = &span
}

// SetColSpan sets column span.
func (c *tableCellImpl) SetColSpan(span int) {
	if c.cell == nil {
		return
	}
	if span < 1 {
		span = 1
	}
	if c.cell.TcPr == nil {
		c.cell.TcPr = &dml.TcPr{}
	}
	c.cell.TcPr.GridSpan = &span
}

func tableFromGraphicFrame(gf *pml.GraphicFrame) *tableImpl {
	if gf == nil || gf.Graphic == nil || gf.Graphic.GraphicData == nil || gf.Graphic.GraphicData.Tbl == nil {
		return nil
	}
	return &tableImpl{tbl: gf.Graphic.GraphicData.Tbl}
}
