package spreadsheet

import (
	"fmt"
	"strconv"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Worksheet represents a worksheet in a workbook.
type Worksheet struct {
	workbook  *Workbook
	worksheet *sml.Worksheet
	name      string
	sheetID   int
	relID     string
	index     int
	path      string
}

// Name returns the worksheet name.
func (ws *Worksheet) Name() string {
	return ws.name
}

// SetName sets the worksheet name.
func (ws *Worksheet) SetName(name string) error {
	ws.name = name
	// Update in workbook reference
	if ws.index < len(ws.workbook.workbook.Sheets.Sheet) {
		ws.workbook.workbook.Sheets.Sheet[ws.index].Name = name
	}
	return nil
}

// Index returns the 0-based sheet index.
func (ws *Worksheet) Index() int {
	return ws.index
}

// Visible returns whether the sheet is visible.
func (ws *Worksheet) Visible() bool {
	if ws.index < len(ws.workbook.workbook.Sheets.Sheet) {
		state := ws.workbook.workbook.Sheets.Sheet[ws.index].State
		return state == "" || state == "visible"
	}
	return true
}

// SetVisible sets whether the sheet is visible.
func (ws *Worksheet) SetVisible(v bool) {
	if ws.index < len(ws.workbook.workbook.Sheets.Sheet) {
		if v {
			ws.workbook.workbook.Sheets.Sheet[ws.index].State = ""
		} else {
			ws.workbook.workbook.Sheets.Sheet[ws.index].State = "hidden"
		}
	}
}

// Hidden returns whether the sheet is hidden.
func (ws *Worksheet) Hidden() bool {
	return !ws.Visible()
}

// SetHidden sets whether the sheet is hidden.
func (ws *Worksheet) SetHidden(v bool) {
	ws.SetVisible(!v)
}

// =============================================================================
// Cell access
// =============================================================================

// Cell returns a cell by reference (e.g., "A1").
func (ws *Worksheet) Cell(ref string) *Cell {
	cellRef, err := utils.ParseCellRef(ref)
	if err != nil {
		return nil
	}
	return ws.CellByRC(cellRef.Row, cellRef.Col)
}

// CellByRC returns a cell by row and column (1-based).
func (ws *Worksheet) CellByRC(row, col int) *Cell {
	if row < 1 || col < 1 {
		return nil
	}

	// Get or create row
	smlRow := ws.getOrCreateRow(row)

	// Get or create cell
	ref := utils.CellRefFromRC(row, col)
	for i, c := range smlRow.C {
		if c.R == ref {
			return &Cell{
				worksheet: ws,
				cell:      smlRow.C[i],
				row:       row,
				col:       col,
			}
		}
	}

	// Create new cell
	cell := &sml.Cell{R: ref}
	smlRow.C = append(smlRow.C, cell)
	return &Cell{
		worksheet: ws,
		cell:      cell,
		row:       row,
		col:       col,
	}
}

// =============================================================================
// Range operations
// =============================================================================

// Range returns a range by reference (e.g., "A1:B5").
func (ws *Worksheet) Range(ref string) *Range {
	rangeRef, err := utils.ParseRangeRef(ref)
	if err != nil {
		return nil
	}
	return &Range{
		worksheet: ws,
		startRow:  rangeRef.Start.Row,
		startCol:  rangeRef.Start.Col,
		endRow:    rangeRef.End.Row,
		endCol:    rangeRef.End.Col,
	}
}

// UsedRange returns the range containing all used cells.
func (ws *Worksheet) UsedRange() *Range {
	maxRow, maxCol := ws.MaxRow(), ws.MaxColumn()
	if maxRow == 0 || maxCol == 0 {
		return nil
	}
	return &Range{
		worksheet: ws,
		startRow:  1,
		startCol:  1,
		endRow:    maxRow,
		endCol:    maxCol,
	}
}

// =============================================================================
// Dimensions
// =============================================================================

// MaxRow returns the highest row number with data.
func (ws *Worksheet) MaxRow() int {
	if ws.worksheet.SheetData == nil {
		return 0
	}
	maxRow := 0
	for _, row := range ws.worksheet.SheetData.Row {
		if row.R > maxRow {
			maxRow = row.R
		}
	}
	return maxRow
}

// MaxColumn returns the highest column number with data.
func (ws *Worksheet) MaxColumn() int {
	if ws.worksheet.SheetData == nil {
		return 0
	}
	maxCol := 0
	for _, row := range ws.worksheet.SheetData.Row {
		for _, cell := range row.C {
			cellRef, err := utils.ParseCellRef(cell.R)
			if err == nil && cellRef.Col > maxCol {
				maxCol = cellRef.Col
			}
		}
	}
	return maxCol
}

// =============================================================================
// Row operations
// =============================================================================

// Row returns a row by index (1-based).
func (ws *Worksheet) Row(index int) *Row {
	smlRow := ws.getOrCreateRow(index)
	return &Row{
		worksheet: ws,
		row:       smlRow,
		index:     index,
	}
}

// =============================================================================
// Merged cells
// =============================================================================

// MergeCells merges a range of cells.
func (ws *Worksheet) MergeCells(ref string) error {
	if ws.worksheet.MergeCells == nil {
		ws.worksheet.MergeCells = &sml.MergeCells{}
	}

	ws.worksheet.MergeCells.MergeCell = append(ws.worksheet.MergeCells.MergeCell, &sml.MergeCell{
		Ref: ref,
	})
	ws.worksheet.MergeCells.Count = len(ws.worksheet.MergeCells.MergeCell)

	return nil
}

// UnmergeCells unmerges a range of cells.
func (ws *Worksheet) UnmergeCells(ref string) error {
	if ws.worksheet.MergeCells == nil {
		return nil
	}

	for i, mc := range ws.worksheet.MergeCells.MergeCell {
		if mc.Ref == ref {
			ws.worksheet.MergeCells.MergeCell = append(
				ws.worksheet.MergeCells.MergeCell[:i],
				ws.worksheet.MergeCells.MergeCell[i+1:]...,
			)
			ws.worksheet.MergeCells.Count = len(ws.worksheet.MergeCells.MergeCell)
			return nil
		}
	}

	return nil
}

// MergedCells returns all merged cell ranges.
func (ws *Worksheet) MergedCells() []string {
	if ws.worksheet.MergeCells == nil {
		return nil
	}

	var refs []string
	for _, mc := range ws.worksheet.MergeCells.MergeCell {
		refs = append(refs, mc.Ref)
	}
	return refs
}

// =============================================================================
// Internal methods
// =============================================================================

func (ws *Worksheet) getOrCreateRow(rowNum int) *sml.Row {
	if ws.worksheet.SheetData == nil {
		ws.worksheet.SheetData = &sml.SheetData{}
	}

	// Find existing row
	for i, row := range ws.worksheet.SheetData.Row {
		if row.R == rowNum {
			return ws.worksheet.SheetData.Row[i]
		}
	}

	// Create new row
	row := &sml.Row{R: rowNum}
	ws.worksheet.SheetData.Row = append(ws.worksheet.SheetData.Row, row)

	// Sort rows by row number
	sortRows(ws.worksheet.SheetData.Row)

	// Find and return the row
	for i, r := range ws.worksheet.SheetData.Row {
		if r.R == rowNum {
			return ws.worksheet.SheetData.Row[i]
		}
	}

	return row
}

func sortRows(rows []*sml.Row) {
	// Simple insertion sort since rows are usually added in order
	for i := 1; i < len(rows); i++ {
		key := rows[i]
		j := i - 1
		for j >= 0 && rows[j].R > key.R {
			rows[j+1] = rows[j]
			j--
		}
		rows[j+1] = key
	}
}

// =============================================================================
// Row type
// =============================================================================

// Row represents a row in a worksheet.
type Row struct {
	worksheet *Worksheet
	row       *sml.Row
	index     int
}

// Index returns the 1-based row index.
func (r *Row) Index() int {
	return r.index
}

// Height returns the row height in points.
func (r *Row) Height() float64 {
	return r.row.Ht
}

// SetHeight sets the row height in points.
func (r *Row) SetHeight(height float64) {
	r.row.Ht = height
	customHeight := true
	r.row.CustomHeight = &customHeight
}

// Hidden returns whether the row is hidden.
func (r *Row) Hidden() bool {
	if r.row.Hidden == nil {
		return false
	}
	return *r.row.Hidden
}

// SetHidden sets whether the row is hidden.
func (r *Row) SetHidden(hidden bool) {
	r.row.Hidden = &hidden
}

// Cell returns a cell in this row by column number (1-based).
func (r *Row) Cell(col int) *Cell {
	return r.worksheet.CellByRC(r.index, col)
}

// Cells returns all cells in this row.
func (r *Row) Cells() []*Cell {
	var cells []*Cell
	for _, c := range r.row.C {
		cellRef, err := utils.ParseCellRef(c.R)
		if err == nil {
			cells = append(cells, &Cell{
				worksheet: r.worksheet,
				cell:      c,
				row:       cellRef.Row,
				col:       cellRef.Col,
			})
		}
	}
	return cells
}

// =============================================================================
// Range type
// =============================================================================

// Range represents a range of cells.
type Range struct {
	worksheet *Worksheet
	startRow  int
	startCol  int
	endRow    int
	endCol    int
}

// Reference returns the A1-style reference (e.g., "A1:B5").
func (r *Range) Reference() string {
	start := utils.CellRefFromRC(r.startRow, r.startCol)
	end := utils.CellRefFromRC(r.endRow, r.endCol)
	return start + ":" + end
}

// RowCount returns the number of rows in the range.
func (r *Range) RowCount() int {
	return r.endRow - r.startRow + 1
}

// ColumnCount returns the number of columns in the range.
func (r *Range) ColumnCount() int {
	return r.endCol - r.startCol + 1
}

// Cells returns all cells in the range as a 2D slice.
func (r *Range) Cells() [][]*Cell {
	rows := make([][]*Cell, r.RowCount())
	for i := 0; i < r.RowCount(); i++ {
		rows[i] = make([]*Cell, r.ColumnCount())
		for j := 0; j < r.ColumnCount(); j++ {
			rows[i][j] = r.worksheet.CellByRC(r.startRow+i, r.startCol+j)
		}
	}
	return rows
}

// ForEach calls fn for each cell in the range.
func (r *Range) ForEach(fn func(cell *Cell) error) error {
	for row := r.startRow; row <= r.endRow; row++ {
		for col := r.startCol; col <= r.endCol; col++ {
			cell := r.worksheet.CellByRC(row, col)
			if err := fn(cell); err != nil {
				return err
			}
		}
	}
	return nil
}

// SetValue sets all cells in the range to the given value.
func (r *Range) SetValue(v interface{}) error {
	return r.ForEach(func(cell *Cell) error {
		return cell.SetValue(v)
	})
}

// Clear clears all cells in the range.
func (r *Range) Clear() error {
	return r.ForEach(func(cell *Cell) error {
		return cell.SetValue(nil)
	})
}

// StartCell returns the top-left cell of the range.
func (r *Range) StartCell() *Cell {
	return r.worksheet.CellByRC(r.startRow, r.startCol)
}

// EndCell returns the bottom-right cell of the range.
func (r *Range) EndCell() *Cell {
	return r.worksheet.CellByRC(r.endRow, r.endCol)
}

// String returns the string representation of the range.
func (r *Range) String() string {
	return fmt.Sprintf("Range[%s]", r.Reference())
}

// =============================================================================
// Helper to convert interface{} to string for cell value
// =============================================================================

func valueToString(v interface{}) string {
	switch val := v.(type) {
	case string:
		return val
	case int:
		return strconv.Itoa(val)
	case int64:
		return strconv.FormatInt(val, 10)
	case float64:
		return strconv.FormatFloat(val, 'f', -1, 64)
	case bool:
		if val {
			return "1"
		}
		return "0"
	default:
		return fmt.Sprintf("%v", v)
	}
}
