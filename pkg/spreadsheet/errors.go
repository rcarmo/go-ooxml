package spreadsheet

import "errors"

var (
	// ErrSheetNotFound is returned when a worksheet is not found.
	ErrSheetNotFound = errors.New("sheet not found")
	// ErrCellNotFound is returned when a cell cannot be resolved.
	ErrCellNotFound  = errors.New("cell not found")
	// ErrInvalidValue is returned when a cell value cannot be set.
	ErrInvalidValue  = errors.New("invalid value")
	// ErrTableNotFound is returned when a table lookup fails.
	ErrTableNotFound = errors.New("table not found")
)
