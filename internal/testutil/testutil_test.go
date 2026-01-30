package testutil

import (
	"testing"
)

func TestHelper(t *testing.T) {
	h := NewHelper(t)
	
	// Test TempFile
	path := h.TempFile("test.txt")
	if path == "" {
		t.Error("TempFile returned empty string")
	}
	
	// Test FileExists (should not exist)
	if h.FileExists(path) {
		t.Error("FileExists should return false for non-existent file")
	}
}

func TestAssertions(t *testing.T) {
	// Create a mock testing.T to capture errors
	// For simplicity, just test that methods don't panic
	h := NewHelper(t)
	
	h.AssertEqual(1, 1, "equal ints")
	h.AssertTrue(true, "true value")
	h.AssertFalse(false, "false value")
	h.AssertContains("hello world", "world", "contains")
}

func TestFormatOptionsString(t *testing.T) {
	tests := []struct {
		opts FormatOptions
		want string
	}{
		{FormatOptions{}, "plain"},
		{FormatOptions{Bold: true}, "bold"},
		{FormatOptions{Italic: true}, "italic"},
		{FormatOptions{Bold: true, Italic: true}, "bold_italic"},
		{FormatOptions{FontSize: 12}, "size"},
		{FormatOptions{Bold: true, FontSize: 14, Color: "FF0000"}, "bold_size_color"},
	}
	
	for _, tc := range tests {
		t.Run(tc.want, func(t *testing.T) {
			got := tc.opts.String()
			if got != tc.want {
				t.Errorf("FormatOptions.String() = %q, want %q", got, tc.want)
			}
		})
	}
}

func TestDocumentType(t *testing.T) {
	tests := []struct {
		dt      DocumentType
		ext     string
		name    string
	}{
		{DocTypeWord, ".docx", "Word"},
		{DocTypePowerPoint, ".pptx", "PowerPoint"},
		{DocTypeExcel, ".xlsx", "Excel"},
	}
	
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			if tc.dt.Extension() != tc.ext {
				t.Errorf("Extension() = %q, want %q", tc.dt.Extension(), tc.ext)
			}
			if tc.dt.String() != tc.name {
				t.Errorf("String() = %q, want %q", tc.dt.String(), tc.name)
			}
		})
	}
}

func TestCommonTestData(t *testing.T) {
	// Verify test data arrays are populated
	if len(CommonStringCases) == 0 {
		t.Error("CommonStringCases is empty")
	}
	if len(CommonNumericCases) == 0 {
		t.Error("CommonNumericCases is empty")
	}
	if len(CommonFormatCombinations) == 0 {
		t.Error("CommonFormatCombinations is empty")
	}
	if len(CommonCellRefCases) == 0 {
		t.Error("CommonCellRefCases is empty")
	}
	if len(CommonRangeCases) == 0 {
		t.Error("CommonRangeCases is empty")
	}
}

type testCloser struct{ closed bool }

func (t *testCloser) Close() error {
	t.closed = true
	return nil
}

func TestResourceHelpers(t *testing.T) {
	t.Run("NewResource", func(t *testing.T) {
		resource := NewResource(t, func() (*testCloser, error) {
			return &testCloser{}, nil
		})
		if resource == nil || resource.closed {
			t.Fatal("NewResource returned invalid resource")
		}
	})

	t.Run("OpenResource", func(t *testing.T) {
		resource := OpenResource(t, func(path string) (*testCloser, error) {
			return &testCloser{}, nil
		}, "dummy")
		if resource == nil || resource.closed {
			t.Fatal("OpenResource returned invalid resource")
		}
	})
}
