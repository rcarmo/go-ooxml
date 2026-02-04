package spreadsheet

import (
	"fmt"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Table represents an Excel table.
type tableImpl struct {
	worksheet *worksheetImpl
	table     *sml.Table
	relID     string
	path      string
}

func newTable(ws *worksheetImpl, id int, name, displayName, ref string) *tableImpl {
	colCount := tableColumnCount(ref)
	columns := make([]*sml.TableColumn, colCount)
	for i := 0; i < colCount; i++ {
		columns[i] = &sml.TableColumn{
			ID:   i + 1,
			Name: fmt.Sprintf("Column%d", i+1),
		}
	}
	table := &sml.Table{
		ID:             id,
		Name:           name,
		DisplayName:    displayName,
		Ref:            ref,
		HeaderRowCount: 1,
		TableColumns: &sml.TableColumns{
			Count:       colCount,
			TableColumn: columns,
		},
		TableStyleInfo: &sml.TableStyleInfo{
			Name:            "TableStyleMedium2",
			ShowRowStripes:  utils.BoolPtr(true),
			ShowColumnStripes: utils.BoolPtr(false),
		},
	}
	return &tableImpl{
		worksheet: ws,
		table:     table,
	}
}

// Name returns the table name.
func (t *tableImpl) Name() string {
	return t.table.Name
}

// DisplayName returns the display name.
func (t *tableImpl) DisplayName() string {
	return t.table.DisplayName
}

// Reference returns the table reference.
func (t *tableImpl) Reference() string {
	return t.table.Ref
}

// Worksheet returns the worksheet owning this table.
func (t *tableImpl) Worksheet() Worksheet {
	return t.worksheet
}

// Headers returns the table column headers.
func (t *tableImpl) Headers() []string {
	if t.table.TableColumns == nil {
		return nil
	}
	headers := make([]string, len(t.table.TableColumns.TableColumn))
	for i, col := range t.table.TableColumns.TableColumn {
		headers[i] = col.Name
	}
	return headers
}

// DataRange returns the data range (excluding header row).
func (t *tableImpl) DataRange() Range {
	start, end, err := parseRange(t.table.Ref)
	if err != nil {
		return nil
	}
	start.Row++
	return &rangeImpl{
		worksheet: t.worksheet,
		startRow:  start.Row,
		startCol:  start.Col,
		endRow:    end.Row,
		endCol:    end.Col,
	}
}

// HasTotalsRow returns true if a totals row is shown.
func (t *tableImpl) HasTotalsRow() bool {
	return utils.DerefBool(t.table.TotalsRowShown, false)
}

// Rows returns the table rows as table rows.
func (t *tableImpl) Rows() []TableRow {
	dataRange := t.DataRange()
	if dataRange == nil {
		return nil
	}
	rowCount := dataRange.RowCount()
	rows := make([]TableRow, rowCount)
	for i := 0; i < rowCount; i++ {
		rows[i] = &tableRowImpl{
			table: t,
			index: i + 1,
		}
	}
	return rows
}

// AddRow appends a row with column values.
func (t *tableImpl) AddRow(values map[string]interface{}) error {
	start, end, err := parseRange(t.table.Ref)
	if err != nil {
		return err
	}
	end.Row++
	t.table.Ref = formatRange(start, end)
	return t.UpdateRow(end.Row-start.Row, values)
}

// UpdateRow updates a row by index (1-based after header).
func (t *tableImpl) UpdateRow(index int, values map[string]interface{}) error {
	if index < 1 {
		return utils.ErrInvalidIndex
	}
	dataRange := t.DataRange()
	if dataRange == nil {
		return utils.ErrInvalidRange
	}
	if index > dataRange.RowCount() {
		return utils.ErrInvalidIndex
	}
	start, _, ok := rangeBounds(dataRange)
	if !ok {
		return utils.ErrInvalidRange
	}
	rowIndex := start.Row + index - 1
	for colName, value := range values {
		colIndex := t.columnIndex(colName)
		if colIndex < 0 {
			continue
		}
		cell := t.worksheet.CellByRC(rowIndex, colIndex+start.Col)
		if cell == nil {
			continue
		}
		if err := cell.SetValue(value); err != nil {
			return err
		}
	}
	return nil
}

// DeleteRow deletes a row by index (1-based after header).
func (t *tableImpl) DeleteRow(index int) error {
	if index < 1 {
		return utils.ErrInvalidIndex
	}
	start, end, err := parseRange(t.table.Ref)
	if err != nil {
		return err
	}
	row := start.Row + index
	if row > end.Row {
		return utils.ErrInvalidIndex
	}
	for col := start.Col; col <= end.Col; col++ {
		cell := t.worksheet.CellByRC(row, col)
		if cell != nil {
			cell.SetValue(nil)
		}
	}
	end.Row--
	if end.Row < start.Row {
		end.Row = start.Row
	}
	t.table.Ref = formatRange(start, end)
	return nil
}

// Column returns all cells in a column by name.
func (t *tableImpl) Column(name string) []Cell {
	dataRange := t.DataRange()
	if dataRange == nil {
		return nil
	}
	colIndex := t.columnIndex(name)
	if colIndex < 0 {
		return nil
	}
	start, _, ok := rangeBounds(dataRange)
	if !ok {
		return nil
	}
	cells := make([]Cell, dataRange.RowCount())
	for i := 0; i < dataRange.RowCount(); i++ {
		cells[i] = t.worksheet.CellByRC(start.Row+i, start.Col+colIndex)
	}
	return cells
}

func (t *tableImpl) columnIndex(name string) int {
	for i, col := range t.table.TableColumns.TableColumn {
		if strings.EqualFold(col.Name, name) {
			return i
		}
	}
	return -1
}

type tableRange struct {
	Row int
	Col int
}

func parseRange(ref string) (tableRange, tableRange, error) {
	rangeRef, err := utils.ParseRangeRef(ref)
	if err != nil {
		return tableRange{}, tableRange{}, err
	}
	return tableRange{Row: rangeRef.Start.Row, Col: rangeRef.Start.Col}, tableRange{Row: rangeRef.End.Row, Col: rangeRef.End.Col}, nil
}

func formatRange(start, end tableRange) string {
	return utils.CellRefFromRC(start.Row, start.Col) + ":" + utils.CellRefFromRC(end.Row, end.Col)
}

func tableColumnCount(ref string) int {
	rangeRef, err := utils.ParseRangeRef(ref)
	if err != nil {
		return 0
	}
	return rangeRef.ColumnCount()
}

// TableRow represents a row in a table.
type tableRowImpl struct {
	table *tableImpl
	index int
}

// Index returns the row index (1-based after header).
func (tr *tableRowImpl) Index() int {
	return tr.index
}

// Values returns the row values keyed by column name.
func (tr *tableRowImpl) Values() map[string]interface{} {
	values := make(map[string]interface{})
	dataRange := tr.table.DataRange()
	if dataRange == nil {
		return values
	}
	start, _, ok := rangeBounds(dataRange)
	if !ok {
		return values
	}
	rowIndex := start.Row + tr.index - 1
	for i, col := range tr.table.table.TableColumns.TableColumn {
		cell := tr.table.worksheet.CellByRC(rowIndex, start.Col+i)
		if cell != nil {
			values[col.Name] = cell.Value()
		}
	}
	return values
}

// Cell returns the cell for a column name.
func (tr *tableRowImpl) Cell(columnName string) Cell {
	dataRange := tr.table.DataRange()
	if dataRange == nil {
		return nil
	}
	colIndex := tr.table.columnIndex(columnName)
	if colIndex < 0 {
		return nil
	}
	start, _, ok := rangeBounds(dataRange)
	if !ok {
		return nil
	}
	rowIndex := start.Row + tr.index - 1
	return tr.table.worksheet.CellByRC(rowIndex, start.Col+colIndex)
}

func rangeBounds(rng Range) (tableRange, tableRange, bool) {
	if rng == nil {
		return tableRange{}, tableRange{}, false
	}
	if r, ok := rng.(*rangeImpl); ok {
		return tableRange{Row: r.startRow, Col: r.startCol}, tableRange{Row: r.endRow, Col: r.endCol}, true
	}
	start, end, err := parseRange(rng.Reference())
	if err != nil {
		return tableRange{}, tableRange{}, false
	}
	return start, end, true
}

// SetValue sets a column value.
func (tr *tableRowImpl) SetValue(columnName string, value interface{}) error {
	cell := tr.Cell(columnName)
	if cell == nil {
		return utils.ErrInvalidIndex
	}
	return cell.SetValue(value)
}
