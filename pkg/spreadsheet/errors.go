package spreadsheet

import "github.com/rcarmo/go-ooxml/pkg/utils"

var (
	// ErrSheetNotFound is returned when a worksheet is not found.
	ErrSheetNotFound = utils.ErrSheetNotFound
	// ErrCellNotFound is returned when a cell cannot be resolved.
	ErrCellNotFound  = utils.ErrCellNotFound
	// ErrInvalidValue is returned when a cell value cannot be set.
	ErrInvalidValue  = utils.ErrInvalidValue
	// ErrTableNotFound is returned when a table lookup fails.
	ErrTableNotFound = utils.ErrTableNotFound
	// ErrCannotDeleteLastSheet is returned when trying to delete the final sheet.
	ErrCannotDeleteLastSheet = utils.ErrCannotDeleteLastSheet
)
