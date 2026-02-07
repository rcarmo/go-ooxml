package spreadsheet

import (
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
)

// Workbook represents an Excel workbook.
type Workbook interface {
	Save() error
	SaveAs(path string) error
	Close() error
	CoreProperties() (*common.CoreProperties, error)
	SetCoreProperties(props *common.CoreProperties) error
	Sheets() []Worksheet
	Sheet(nameOrIndex interface{}) (Worksheet, error)
	AddSheet(name string) Worksheet
	DeleteSheet(nameOrIndex interface{}) error
	NamedRanges() []NamedRange
	AddNamedRange(name, refersTo string) NamedRange
	Tables() []Table
	Table(name string) (Table, error)
	Styles() Styles
	SheetsRaw() []*worksheetImpl
	SheetRaw(nameOrIndex interface{}) (*worksheetImpl, error)
	SheetCount() int
	SharedStrings() *SharedStrings
}

// WorkbookInterface is a compatibility alias for the workbook interface.
type WorkbookInterface = Workbook

// Worksheet represents a worksheet.
type Worksheet interface {
	Name() string
	SetName(name string) error
	Index() int
	Visible() bool
	SetVisible(v bool)
	Hidden() bool
	SetHidden(v bool)
	Cell(ref string) Cell
	CellByRC(row, col int) Cell
	Range(ref string) Range
	Row(index int) Row
	Rows() RowIterator
	UsedRange() Range
	MaxRow() int
	MaxColumn() int
	Tables() []Table
	AddTable(ref string, name string) Table
	MergedCells() []Range
	MergeCells(ref string) error
	UnmergeCells(ref string) error
	AddChart(fromCell, toCell, title string) error
	AddDiagram(fromCell, toCell, title string) error
	AddPicture(imagePath, fromCell, toCell string) error
	Comments() []Comment
	PageMargins() (PageMargins, bool)
	SetPageMargins(margins PageMargins)
}

// WorksheetInterface is a compatibility alias for the worksheet interface.
type WorksheetInterface = Worksheet

// Cell represents a cell.
type Cell interface {
	Reference() string
	Row() int
	Column() int
	Value() interface{}
	SetValue(v interface{}) error
	String() string
	Float64() (float64, error)
	Int() (int, error)
	Bool() (bool, error)
	Time() (time.Time, error)
	Formula() string
	SetFormula(formula string) error
	HasFormula() bool
	Type() CellType
	Style() CellStyle
	SetStyle(style CellStyle) error
	NumberFormat() string
	SetNumberFormat(format string) error
	Comment() (Comment, bool)
	SetComment(text, author string) error
}

// CellInterface is a compatibility alias for the cell interface.
type CellInterface = Cell

// Range represents a range of cells.
type Range interface {
	Reference() string
	StartCell() Cell
	EndCell() Cell
	Cells() [][]Cell
	ForEach(fn func(cell Cell) error) error
	SetValue(v interface{}) error
	Clear() error
	RowCount() int
	ColumnCount() int
}

// RangeInterface is a compatibility alias for the range interface.
type RangeInterface = Range

// Table represents an Excel table.
type Table interface {
	Name() string
	DisplayName() string
	Reference() string
	Worksheet() Worksheet
	Headers() []string
	DataRange() Range
	HasTotalsRow() bool
	Rows() []TableRow
	AddRow(values map[string]interface{}) error
	UpdateRow(index int, values map[string]interface{}) error
	DeleteRow(index int) error
	Column(name string) []Cell
}

// TableInterface is a compatibility alias for the table interface.
type TableInterface = Table

// TableRow represents a table row (1-based index after header).
type TableRow interface {
	Index() int
	Values() map[string]interface{}
	Cell(columnName string) Cell
	SetValue(columnName string, value interface{}) error
}

// TableRowInterface is a compatibility alias for the table row interface.
type TableRowInterface = TableRow

// Row represents a worksheet row.
type Row interface {
	Index() int
	Height() float64
	SetHeight(height float64)
	Hidden() bool
	SetHidden(hidden bool)
	Cell(col int) Cell
	Cells() []Cell
}

// RowInterface is a compatibility alias for the row interface.
type RowInterface = Row

// RowIterator iterates rows in a worksheet.
type RowIterator interface {
	Next() (Row, bool)
}

// RowIteratorInterface is a compatibility alias for the row iterator interface.
type RowIteratorInterface = RowIterator

// NamedRange represents a defined name in a workbook.
type NamedRange interface {
	Name() string
	RefersTo() string
	SetRefersTo(refersTo string)
	SheetIndex() (int, bool)
	SetSheetIndex(index int)
	ClearSheetIndex()
	Hidden() bool
	SetHidden(hidden bool)
}

// NamedRangeInterface is a compatibility alias for the named range interface.
type NamedRangeInterface = NamedRange

// Comment represents a worksheet comment.
type Comment interface {
	Reference() string
	Author() string
	Text() string
	SetText(text string)
}

// CommentInterface is a compatibility alias for the comment interface.
type CommentInterface = Comment

// Styles represents workbook styles.
type Styles interface {
	Style() CellStyle
}

// StylesInterface is a compatibility alias for the styles interface.
type StylesInterface = Styles

// CellStyle represents cell formatting.
type CellStyle interface {
	FontName() string
	SetFontName(name string) CellStyle
	FontSize() float64
	SetFontSize(size float64) CellStyle
	Bold() bool
	SetBold(v bool) CellStyle
	Italic() bool
	SetItalic(v bool) CellStyle
	FillColor() string
	SetFillColor(hex string) CellStyle
	Border() Border
	SetBorder(border Border) CellStyle
	HorizontalAlignment() Alignment
	SetHorizontalAlignment(a Alignment) CellStyle
	VerticalAlignment() Alignment
	SetVerticalAlignment(a Alignment) CellStyle
	NumberFormat() string
	SetNumberFormat(format string) CellStyle
}

// PageMargins represents worksheet page margins.
type PageMargins = sml.PageMargins

// CellStyleInterface is a compatibility alias for the cell style interface.
type CellStyleInterface = CellStyle
