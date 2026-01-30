package utils

import (
	"testing"
)

func TestValidationError(t *testing.T) {
	tests := []struct {
		name    string
		field   string
		message string
		value   interface{}
		want    string
	}{
		{
			name:    "with value",
			field:   "cellRef",
			message: "invalid format",
			value:   "ZZZ999999",
			want:    "cellRef: invalid format (value: ZZZ999999)",
		},
		{
			name:    "without value",
			field:   "range",
			message: "cannot be empty",
			value:   nil,
			want:    "range: cannot be empty",
		},
		{
			name:    "numeric value",
			field:   "row",
			message: "must be positive",
			value:   -1,
			want:    "row: must be positive (value: -1)",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := NewValidationError(tt.field, tt.message, tt.value)
			if got := err.Error(); got != tt.want {
				t.Errorf("ValidationError.Error() = %q, want %q", got, tt.want)
			}
			if err.Field != tt.field {
				t.Errorf("Field = %q, want %q", err.Field, tt.field)
			}
			if err.Message != tt.message {
				t.Errorf("Message = %q, want %q", err.Message, tt.message)
			}
		})
	}
}

func TestStandardErrors(t *testing.T) {
	// Ensure all standard errors are non-nil and have messages
	errors := []struct {
		name string
		err  error
	}{
		{"ErrDocumentClosed", ErrDocumentClosed},
		{"ErrPartNotFound", ErrPartNotFound},
		{"ErrInvalidCellRef", ErrInvalidCellRef},
		{"ErrInvalidRange", ErrInvalidRange},
		{"ErrTableNotFound", ErrTableNotFound},
		{"ErrSectionNotFound", ErrSectionNotFound},
		{"ErrSlideNotFound", ErrSlideNotFound},
		{"ErrShapeNotFound", ErrShapeNotFound},
		{"ErrInvalidIndex", ErrInvalidIndex},
		{"ErrReadOnly", ErrReadOnly},
		{"ErrInvalidFormat", ErrInvalidFormat},
		{"ErrCorruptedFile", ErrCorruptedFile},
	}

	for _, tt := range errors {
		t.Run(tt.name, func(t *testing.T) {
			if tt.err == nil {
				t.Error("error is nil")
			}
			if tt.err.Error() == "" {
				t.Error("error message is empty")
			}
		})
	}
}
