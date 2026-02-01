package spreadsheet

import "errors"

// Errors
var (
	ErrSheetNotFound = errors.New("sheet not found")
	ErrCellNotFound  = errors.New("cell not found")
	ErrInvalidValue  = errors.New("invalid value")
	ErrTableNotFound = errors.New("table not found")
)
