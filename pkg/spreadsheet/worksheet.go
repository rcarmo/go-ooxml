package spreadsheet

import (
	"fmt"
	"strconv"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Worksheet represents a worksheet in a workbook.
type worksheetImpl struct {
	workbook  *workbookImpl
	worksheet *sml.Worksheet
	name      string
	sheetID   int
	relID     string
	index     int
	path      string
	tables    []*tableImpl
	comments  *SheetComments
}

// Name returns the worksheet name.
func (ws *worksheetImpl) Name() string {
	return ws.name
}

// SetName sets the worksheet name.
func (ws *worksheetImpl) SetName(name string) error {
	ws.name = name
	// Update in workbook reference
	if ws.index < len(ws.workbook.workbook.Sheets.Sheet) {
		ws.workbook.workbook.Sheets.Sheet[ws.index].Name = name
	}
	return nil
}

// Index returns the 0-based sheet index.
func (ws *worksheetImpl) Index() int {
	return ws.index
}

// Visible returns whether the sheet is visible.
func (ws *worksheetImpl) Visible() bool {
	if ws.index < len(ws.workbook.workbook.Sheets.Sheet) {
		state := ws.workbook.workbook.Sheets.Sheet[ws.index].State
		return state == "" || state == "visible"
	}
	return true
}

// SetVisible sets whether the sheet is visible.
func (ws *worksheetImpl) SetVisible(v bool) {
	if ws.index < len(ws.workbook.workbook.Sheets.Sheet) {
		if v {
			ws.workbook.workbook.Sheets.Sheet[ws.index].State = ""
		} else {
			ws.workbook.workbook.Sheets.Sheet[ws.index].State = "hidden"
		}
	}
}

// Hidden returns whether the sheet is hidden.
func (ws *worksheetImpl) Hidden() bool {
	return !ws.Visible()
}

// SetHidden sets whether the sheet is hidden.
func (ws *worksheetImpl) SetHidden(v bool) {
	ws.SetVisible(!v)
}

// =============================================================================
// Table access
// =============================================================================

// Tables returns all tables in the worksheet.
func (ws *worksheetImpl) Tables() []Table {
	result := make([]Table, len(ws.tables))
	for i, table := range ws.tables {
		result[i] = table
	}
	return result
}

// AddTable adds a new table to the worksheet.
func (ws *worksheetImpl) AddTable(ref string, name string) Table {
	tableID := ws.workbook.nextTableID
	ws.workbook.nextTableID++
	if name == "" {
		name = fmt.Sprintf("Table%d", tableID)
	}
	displayName := name

	table := newTable(ws, tableID, name, displayName, ref)
	table.relID = ws.nextTableRelID()
	ws.tables = append(ws.tables, table)
	ws.addTablePart(table.relID)
	ws.writeTableHeaders(table)
	return table
}

// Table returns a table by name.
func (ws *worksheetImpl) Table(name string) (Table, error) {
	for _, table := range ws.tables {
		if table.Name() == name {
			return table, nil
		}
	}
	return nil, ErrTableNotFound
}

// =============================================================================
// Cell access
// =============================================================================

// Cell returns a cell by reference (e.g., "A1").
func (ws *worksheetImpl) Cell(ref string) Cell {
	cellRef, err := utils.ParseCellRef(ref)
	if err != nil {
		return nil
	}
	return ws.CellByRC(cellRef.Row, cellRef.Col)
}

// CellByRC returns a cell by row and column (1-based).
func (ws *worksheetImpl) CellByRC(row, col int) Cell {
	if row < 1 || col < 1 {
		return nil
	}

	// Get or create row
	smlRow := ws.getOrCreateRow(row)

	// Get or create cell
	ref := utils.CellRefFromRC(row, col)
	for i, c := range smlRow.C {
		if c.R == ref {
			return &cellImpl{
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
	return &cellImpl{
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
func (ws *worksheetImpl) Range(ref string) Range {
	rangeRef, err := utils.ParseRangeRef(ref)
	if err != nil {
		return nil
	}
	return &rangeImpl{
			worksheet: ws,
			startRow:  rangeRef.Start.Row,
			startCol:  rangeRef.Start.Col,
			endRow:    rangeRef.End.Row,
			endCol:    rangeRef.End.Col,
		}
}

// UsedRange returns the range containing all used cells.
func (ws *worksheetImpl) UsedRange() Range {
	maxRow, maxCol := ws.MaxRow(), ws.MaxColumn()
	if maxRow == 0 || maxCol == 0 {
		return nil
	}
	return &rangeImpl{
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
func (ws *worksheetImpl) MaxRow() int {
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
func (ws *worksheetImpl) MaxColumn() int {
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
func (ws *worksheetImpl) Row(index int) Row {
	smlRow := ws.getOrCreateRow(index)
	return &rowImpl{
		worksheet: ws,
		row:       smlRow,
		index:     index,
	}
}

// Rows returns an iterator over worksheet rows.
func (ws *worksheetImpl) Rows() RowIterator {
	return &rowIterator{
		worksheet: ws,
		index:     0,
	}
}

// =============================================================================
// Merged cells
// =============================================================================

// MergeCells merges a range of cells.
func (ws *worksheetImpl) MergeCells(ref string) error {
	if ws.worksheet.MergeCells == nil {
		ws.worksheet.MergeCells = &sml.MergeCells{}
	}

	ws.worksheet.MergeCells.MergeCell = append(ws.worksheet.MergeCells.MergeCell, &sml.MergeCell{
		Ref: ref,
	})
	ws.worksheet.MergeCells.Count = len(ws.worksheet.MergeCells.MergeCell)

	return nil
}

// AddChart adds a chart to the worksheet.
func (ws *worksheetImpl) AddChart(fromCell, toCell, title string) error {
	return ws.addGraphic(fromCell, toCell, title, drawingKindChart, "")
}

// AddDiagram adds a diagram (SmartArt) to the worksheet.
func (ws *worksheetImpl) AddDiagram(fromCell, toCell, title string) error {
	return ws.addGraphic(fromCell, toCell, title, drawingKindDiagram, "")
}

// AddPicture adds an image to the worksheet.
func (ws *worksheetImpl) AddPicture(imagePath, fromCell, toCell string) error {
	return ws.addGraphic(fromCell, toCell, "", drawingKindPicture, imagePath)
}

// UnmergeCells unmerges a range of cells.
func (ws *worksheetImpl) UnmergeCells(ref string) error {
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
func (ws *worksheetImpl) MergedCells() []Range {
	if ws.worksheet.MergeCells == nil {
		return nil
	}

	var refs []Range
	for _, mc := range ws.worksheet.MergeCells.MergeCell {
		rng := ws.Range(mc.Ref)
		if rng != nil {
			refs = append(refs, rng)
		}
	}
	return refs
}

// =============================================================================
// Internal methods
// =============================================================================

func (ws *worksheetImpl) getOrCreateRow(rowNum int) *sml.Row {
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

func (ws *worksheetImpl) addTablePart(relID string) {
	if ws.worksheet.TableParts == nil {
		ws.worksheet.TableParts = &sml.TableParts{}
	}
	ws.worksheet.TableParts.TablePart = append(ws.worksheet.TableParts.TablePart, &sml.TablePart{ID: relID})
	ws.worksheet.TableParts.Count = len(ws.worksheet.TableParts.TablePart)
}

func (ws *worksheetImpl) nextTableRelID() string {
	sheetPath := ws.path
	if sheetPath == "" {
		sheetPath = fmt.Sprintf("xl/worksheets/sheet%d.xml", ws.index+1)
	}
	rels := ws.workbook.pkg.GetRelationships(sheetPath)
	return rels.NextID()
}

func (ws *worksheetImpl) writeTableHeaders(table *tableImpl) {
	start, _, err := parseRange(table.table.Ref)
	if err != nil {
		return
	}
	for i, header := range table.Headers() {
		cell := ws.CellByRC(start.Row, start.Col+i)
		if cell != nil {
			_ = cell.SetValue(header)
		}
	}
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
type rowImpl struct {
	worksheet *worksheetImpl
	row       *sml.Row
	index     int
}

// Index returns the 1-based row index.
func (r *rowImpl) Index() int {
	return r.index
}

// Height returns the row height in points.
func (r *rowImpl) Height() float64 {
	return r.row.Ht
}

// SetHeight sets the row height in points.
func (r *rowImpl) SetHeight(height float64) {
	r.row.Ht = height
	customHeight := true
	r.row.CustomHeight = &customHeight
}

// Hidden returns whether the row is hidden.
func (r *rowImpl) Hidden() bool {
	if r.row.Hidden == nil {
		return false
	}
	return *r.row.Hidden
}

// SetHidden sets whether the row is hidden.
func (r *rowImpl) SetHidden(hidden bool) {
	r.row.Hidden = &hidden
}

// Cell returns a cell in this row by column number (1-based).
func (r *rowImpl) Cell(col int) Cell {
	return r.worksheet.CellByRC(r.index, col)
}

// Cells returns all cells in this row.
func (r *rowImpl) Cells() []Cell {
	var cells []Cell
	for _, c := range r.row.C {
		cellRef, err := utils.ParseCellRef(c.R)
		if err == nil {
			cells = append(cells, &cellImpl{
				worksheet: r.worksheet,
				cell:      c,
				row:       cellRef.Row,
				col:       cellRef.Col,
			})
		}
	}
	return cells
}

type rowIterator struct {
	worksheet *worksheetImpl
	index     int
}

// Next advances the iterator and returns the next row.
func (it *rowIterator) Next() (Row, bool) {
	if it == nil || it.worksheet == nil || it.worksheet.worksheet == nil || it.worksheet.worksheet.SheetData == nil {
		return nil, false
	}
	if it.index >= len(it.worksheet.worksheet.SheetData.Row) {
		return nil, false
	}
	smlRow := it.worksheet.worksheet.SheetData.Row[it.index]
	rowIndex := smlRow.R
	if rowIndex == 0 {
		rowIndex = it.index + 1
	}
	it.index++
	return &rowImpl{
		worksheet: it.worksheet,
		row:       smlRow,
		index:     rowIndex,
	}, true
}

// =============================================================================
// Range type
// =============================================================================

// rangeImpl represents a range of cells.
type rangeImpl struct {
	worksheet *worksheetImpl
	startRow  int
	startCol  int
	endRow    int
	endCol    int
}

// Reference returns the A1-style reference (e.g., "A1:B5").
func (r *rangeImpl) Reference() string {
	start := utils.CellRefFromRC(r.startRow, r.startCol)
	end := utils.CellRefFromRC(r.endRow, r.endCol)
	return start + ":" + end
}

// RowCount returns the number of rows in the range.
func (r *rangeImpl) RowCount() int {
	return r.endRow - r.startRow + 1
}

// ColumnCount returns the number of columns in the range.
func (r *rangeImpl) ColumnCount() int {
	return r.endCol - r.startCol + 1
}

// Cells returns all cells in the range as a 2D slice.
func (r *rangeImpl) Cells() [][]Cell {
	rows := make([][]Cell, r.RowCount())
	for i := 0; i < r.RowCount(); i++ {
		rows[i] = make([]Cell, r.ColumnCount())
		for j := 0; j < r.ColumnCount(); j++ {
			rows[i][j] = r.worksheet.CellByRC(r.startRow+i, r.startCol+j)
		}
	}
	return rows
}

// ForEach calls fn for each cell in the range.
func (r *rangeImpl) ForEach(fn func(cell Cell) error) error {
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
func (r *rangeImpl) SetValue(v interface{}) error {
	return r.ForEach(func(cell Cell) error {
		return cell.SetValue(v)
	})
}

// Clear clears all cells in the range.
func (r *rangeImpl) Clear() error {
	return r.ForEach(func(cell Cell) error {
		return cell.SetValue(nil)
	})
}

// StartCell returns the top-left cell of the range.
func (r *rangeImpl) StartCell() Cell {
	return r.worksheet.CellByRC(r.startRow, r.startCol)
}

// EndCell returns the bottom-right cell of the range.
func (r *rangeImpl) EndCell() Cell {
	return r.worksheet.CellByRC(r.endRow, r.endCol)
}

// String returns the string representation of the range.
func (r *rangeImpl) String() string {
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
