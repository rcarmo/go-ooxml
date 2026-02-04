package document

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Index returns the table index within the body.
func (t *tableImpl) Index() int {
	return t.index
}

// Rows returns all rows in the table.
func (t *tableImpl) Rows() []Row {
	result := make([]Row, len(t.tbl.Tr))
	for i, tr := range t.tbl.Tr {
		result[i] = &rowImpl{doc: t.doc, tr: tr, index: i}
	}
	return result
}

// Row returns the row at the given index (0-based).
func (t *tableImpl) Row(index int) Row {
	if index < 0 || index >= len(t.tbl.Tr) {
		return nil
	}
	return &rowImpl{doc: t.doc, tr: t.tbl.Tr[index], index: index}
}

// AddRow adds a new row at the end of the table.
func (t *tableImpl) AddRow() Row {
	colCount := t.ColumnCount()
	tr := &wml.Tr{}
	for i := 0; i < colCount; i++ {
		tc := &wml.Tc{
			Content: []interface{}{&wml.P{}},
		}
		tr.Tc = append(tr.Tc, tc)
	}
	t.tbl.Tr = append(t.tbl.Tr, tr)
	return &rowImpl{doc: t.doc, tr: tr, index: len(t.tbl.Tr) - 1}
}

// Purpose returns the inferred table purpose based on headers.
func (t *tableImpl) Purpose() string {
	headers := t.FirstRowText()
	if len(headers) == 0 {
		return ""
	}
	return strings.Join(headers, ", ")
}

// InsertRow inserts a new row at the given index.
func (t *tableImpl) InsertRow(index int) Row {
	colCount := t.ColumnCount()
	tr := &wml.Tr{}
	for i := 0; i < colCount; i++ {
		tc := &wml.Tc{
			Content: []interface{}{&wml.P{}},
		}
		tr.Tc = append(tr.Tc, tc)
	}

	if index >= len(t.tbl.Tr) {
		t.tbl.Tr = append(t.tbl.Tr, tr)
	} else {
		t.tbl.Tr = append(t.tbl.Tr[:index+1], t.tbl.Tr[index:]...)
		t.tbl.Tr[index] = tr
	}

	return &rowImpl{doc: t.doc, tr: tr, index: index}
}

// DeleteRow deletes the row at the given index.
func (t *tableImpl) DeleteRow(index int) error {
	if index < 0 || index >= len(t.tbl.Tr) {
		return utils.ErrInvalidIndex
	}
	t.tbl.Tr = append(t.tbl.Tr[:index], t.tbl.Tr[index+1:]...)
	return nil
}

// Cell returns the cell at the given row and column (0-based).
func (t *tableImpl) Cell(row, col int) Cell {
	r := t.Row(row)
	if r == nil {
		return nil
	}
	return r.Cell(col)
}

// RowCount returns the number of rows.
func (t *tableImpl) RowCount() int {
	return len(t.tbl.Tr)
}

// ColumnCount returns the number of columns (based on grid or first row).
func (t *tableImpl) ColumnCount() int {
	if t.tbl.TblGrid != nil && len(t.tbl.TblGrid.GridCol) > 0 {
		return len(t.tbl.TblGrid.GridCol)
	}
	if len(t.tbl.Tr) > 0 {
		return len(t.tbl.Tr[0].Tc)
	}
	return 0
}

// FirstRowText returns the text of all cells in the first row.
func (t *tableImpl) FirstRowText() []string {
	if len(t.tbl.Tr) == 0 {
		return nil
	}
	row := t.Row(0)
	cells := row.Cells()
	result := make([]string, len(cells))
	for i, cell := range cells {
		result[i] = cell.Text()
	}
	return result
}

// Style returns the table style ID.
func (t *tableImpl) Style() string {
	if t.tbl.TblPr != nil && t.tbl.TblPr.TblStyle != nil {
		return t.tbl.TblPr.TblStyle.Val
	}
	return ""
}

// SetStyle sets the table style.
func (t *tableImpl) SetStyle(styleID string) {
	if t.tbl.TblPr == nil {
		t.tbl.TblPr = &wml.TblPr{}
	}
	t.tbl.TblPr.TblStyle = &wml.TblStyle{Val: styleID}
}

// XML returns the underlying WML table for advanced access.
func (t *tableImpl) XML() *wml.Tbl {
	return t.tbl
}

// Cells returns all cells in the row.
func (r *rowImpl) Cells() []Cell {
	result := make([]Cell, len(r.tr.Tc))
	for i, tc := range r.tr.Tc {
		result[i] = &cellImpl{doc: r.doc, tc: tc, index: i}
	}
	return result
}

// Cell returns the cell at the given index (0-based).
func (r *rowImpl) Cell(index int) Cell {
	if index < 0 || index >= len(r.tr.Tc) {
		return nil
	}
	return &cellImpl{doc: r.doc, tc: r.tr.Tc[index], index: index}
}

// AddCell adds a new cell at the end of the row.
func (r *rowImpl) AddCell() Cell {
	tc := &wml.Tc{
		Content: []interface{}{&wml.P{}},
	}
	r.tr.Tc = append(r.tr.Tc, tc)
	return &cellImpl{doc: r.doc, tc: tc, index: len(r.tr.Tc) - 1}
}

// Index returns the row's index in the table.
func (r *rowImpl) Index() int {
	return r.index
}

// IsHeader returns whether this row is a header row.
func (r *rowImpl) IsHeader() bool {
	if r.tr.TrPr != nil && r.tr.TrPr.TblHeader != nil {
		return r.tr.TrPr.TblHeader.Enabled()
	}
	return false
}

// SetHeader sets whether this row is a header row.
func (r *rowImpl) SetHeader(v bool) {
	if r.tr.TrPr == nil {
		r.tr.TrPr = &wml.TrPr{}
	}
	if v {
		r.tr.TrPr.TblHeader = wml.NewOnOffEnabled()
	} else {
		r.tr.TrPr.TblHeader = nil
	}
}

// XML returns the underlying WML row for advanced access.
func (r *rowImpl) XML() *wml.Tr {
	return r.tr
}

// Text returns the combined text of all paragraphs in the cell.
func (c *cellImpl) Text() string {
	var sb strings.Builder
	for i, para := range c.Paragraphs() {
		if i > 0 {
			sb.WriteString("\n")
		}
		sb.WriteString(para.Text())
	}
	return sb.String()
}

// SetText sets the cell text, replacing all existing content.
func (c *cellImpl) SetText(text string) {
	c.tc.Content = []interface{}{&wml.P{}}
	if para := c.firstParagraph(); para != nil {
		para.SetText(text)
	}
}

// Paragraphs returns all paragraphs in the cell.
func (c *cellImpl) Paragraphs() []Paragraph {
	var result []Paragraph
	for i, elem := range c.tc.Content {
		if p, ok := elem.(*wml.P); ok {
			result = append(result, &paragraphImpl{doc: c.doc, p: p, index: i})
		}
	}
	return result
}

// AddParagraph adds a new paragraph to the cell.
func (c *cellImpl) AddParagraph() Paragraph {
	p := &wml.P{}
	c.tc.Content = append(c.tc.Content, p)
	return &paragraphImpl{doc: c.doc, p: p, index: len(c.tc.Content) - 1}
}

func (c *cellImpl) firstParagraph() *paragraphImpl {
	paras := c.Paragraphs()
	if len(paras) > 0 {
		if p, ok := paras[0].(*paragraphImpl); ok {
			return p
		}
	}
	if p, ok := c.AddParagraph().(*paragraphImpl); ok {
		return p
	}
	return nil
}

// GridSpan returns the column span.
func (c *cellImpl) GridSpan() int {
	if c.tc.TcPr != nil && c.tc.TcPr.GridSpan != nil {
		return c.tc.TcPr.GridSpan.Val
	}
	return 1
}

// SetGridSpan sets the column span.
func (c *cellImpl) SetGridSpan(span int) {
	if c.tc.TcPr == nil {
		c.tc.TcPr = &wml.TcPr{}
	}
	if span <= 1 {
		c.tc.TcPr.GridSpan = nil
	} else {
		c.tc.TcPr.GridSpan = &wml.GridSpan{Val: span}
	}
}

// VerticalMerge returns the vertical merge type.
func (c *cellImpl) VerticalMerge() VerticalMerge {
	if c.tc.TcPr != nil && c.tc.TcPr.VMerge != nil {
		return VerticalMerge(c.tc.TcPr.VMerge.Val)
	}
	return ""
}

// SetVerticalMerge sets the vertical merge type ("restart" or "continue").
func (c *cellImpl) SetVerticalMerge(val VerticalMerge) {
	if c.tc.TcPr == nil {
		c.tc.TcPr = &wml.TcPr{}
	}
	if val == "" {
		c.tc.TcPr.VMerge = nil
	} else {
		c.tc.TcPr.VMerge = &wml.VMerge{Val: string(val)}
	}
}

// Shading returns the cell background color.
func (c *cellImpl) Shading() string {
	if c.tc.TcPr != nil && c.tc.TcPr.Shd != nil {
		return c.tc.TcPr.Shd.Fill
	}
	return ""
}

// SetShading sets the cell background color (hex without #).
func (c *cellImpl) SetShading(fill string) {
	if c.tc.TcPr == nil {
		c.tc.TcPr = &wml.TcPr{}
	}
	c.tc.TcPr.Shd = &wml.Shd{
		Val:  "clear",
		Fill: strings.TrimPrefix(fill, "#"),
	}
}

// Index returns the cell's index in the row.
func (c *cellImpl) Index() int {
	return c.index
}

// XML returns the underlying WML cell for advanced access.
func (c *cellImpl) XML() *wml.Tc {
	return c.tc
}
