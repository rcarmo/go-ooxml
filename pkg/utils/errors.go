// Package utils provides shared utilities for OOXML manipulation.
package utils

import (
	"errors"
	"fmt"
)

var (
	// ErrDocumentClosed is returned when operating on a closed document.
	ErrDocumentClosed  = errors.New("document is closed")
	// ErrPartNotFound is returned when a package part is missing.
	ErrPartNotFound    = errors.New("part not found")
	// ErrInvalidCellRef is returned for invalid cell references.
	ErrInvalidCellRef  = errors.New("invalid cell reference")
	// ErrInvalidRange is returned for invalid range references.
	ErrInvalidRange    = errors.New("invalid range")
	// ErrTableNotFound is returned when a table lookup fails.
	ErrTableNotFound   = errors.New("table not found")
	// ErrSectionNotFound is returned when a section cannot be found.
	ErrSectionNotFound = errors.New("section not found")
	// ErrSlideNotFound is returned when a slide lookup fails.
	ErrSlideNotFound   = errors.New("slide not found")
	// ErrShapeNotFound is returned when a shape lookup fails.
	ErrShapeNotFound   = errors.New("shape not found")
	// ErrInvalidIndex is returned when an index is out of range.
	ErrInvalidIndex    = errors.New("invalid index")
	// ErrReadOnly is returned when trying to modify a read-only document.
	ErrReadOnly        = errors.New("document is read-only")
	// ErrInvalidFormat is returned for unsupported file formats.
	ErrInvalidFormat   = errors.New("invalid document format")
	// ErrCorruptedFile is returned when the file cannot be parsed.
	ErrCorruptedFile   = errors.New("corrupted file")
)

// ValidationError provides detailed validation failure info.
type ValidationError struct {
	Field   string
	Message string
	Value   interface{}
}

// Error returns a formatted validation error string.
func (e *ValidationError) Error() string {
	if e.Value != nil {
		return fmt.Sprintf("%s: %s (value: %v)", e.Field, e.Message, e.Value)
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
