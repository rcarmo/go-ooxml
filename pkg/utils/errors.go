// Package utils provides shared utilities for OOXML manipulation.
package utils

import "errors"

// Standard errors for the library
var (
	ErrDocumentClosed  = errors.New("document is closed")
	ErrPartNotFound    = errors.New("part not found")
	ErrInvalidCellRef  = errors.New("invalid cell reference")
	ErrInvalidRange    = errors.New("invalid range")
	ErrTableNotFound   = errors.New("table not found")
	ErrSectionNotFound = errors.New("section not found")
	ErrSlideNotFound   = errors.New("slide not found")
	ErrShapeNotFound   = errors.New("shape not found")
	ErrInvalidIndex    = errors.New("invalid index")
	ErrReadOnly        = errors.New("document is read-only")
	ErrInvalidFormat   = errors.New("invalid document format")
	ErrCorruptedFile   = errors.New("corrupted file")
)

// ValidationError provides detailed validation failure info.
type ValidationError struct {
	Field   string
	Message string
	Value   interface{}
}

func (e *ValidationError) Error() string {
	if e.Value != nil {
		return e.Field + ": " + e.Message
	}
	return e.Field + ": " + e.Message
}

// NewValidationError creates a new validation error.
func NewValidationError(field, message string, value interface{}) *ValidationError {
	return &ValidationError{
		Field:   field,
		Message: message,
		Value:   value,
	}
}
