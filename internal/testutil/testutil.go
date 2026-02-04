// Package testutil provides shared testing utilities across all OOXML packages.
package testutil

import (
	"os"
	"path/filepath"
	"testing"
)

// =============================================================================
// Test Helper - Generic test lifecycle management
// =============================================================================

// Helper provides common test utilities.
type Helper struct {
	T       *testing.T
	TempDir string
}

// NewHelper creates a new test helper.
func NewHelper(t *testing.T) *Helper {
	t.Helper()
	return &Helper{
		T:       t,
		TempDir: t.TempDir(),
	}
}

// Closer represents a resource that can be closed.
type Closer interface {
	Close() error
}

// NewResource creates a new resource and registers cleanup.
func NewResource[T Closer](t *testing.T, create func() (T, error)) T {
	t.Helper()
	resource, err := create()
	if err != nil {
		t.Fatalf("create resource: %v", err)
	}
	t.Cleanup(func() {
		_ = resource.Close()
	})
	return resource
}

// OpenResource opens an existing resource and registers cleanup.
func OpenResource[T Closer](t *testing.T, open func(string) (T, error), path string) T {
	t.Helper()
	resource, err := open(path)
	if err != nil {
		t.Fatalf("open resource: %v", err)
	}
	t.Cleanup(func() {
		_ = resource.Close()
	})
	return resource
}

// TempFile returns a path for a temporary file with the given name.
func (h *Helper) TempFile(name string) string {
	return filepath.Join(h.TempDir, name)
}

// FileExists checks if a file exists.
func (h *Helper) FileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}

// RequireNoError fails the test immediately if err is not nil.
func (h *Helper) RequireNoError(err error, msg string) {
	h.T.Helper()
	if err != nil {
		h.T.Fatalf("%s: %v", msg, err)
	}
}

// AssertEqual checks if got equals want.
func (h *Helper) AssertEqual(got, want interface{}, msg string) {
	h.T.Helper()
	if got != want {
		h.T.Errorf("%s: got %v, want %v", msg, got, want)
	}
}

// AssertNotNil checks if v is not nil.
func (h *Helper) AssertNotNil(v interface{}, msg string) {
	h.T.Helper()
	if v == nil {
		h.T.Errorf("%s: expected non-nil", msg)
	}
}

// AssertTrue checks if v is true.
func (h *Helper) AssertTrue(v bool, msg string) {
	h.T.Helper()
	if !v {
		h.T.Errorf("%s: expected true", msg)
	}
}

// AssertFalse checks if v is false.
func (h *Helper) AssertFalse(v bool, msg string) {
	h.T.Helper()
	if v {
		h.T.Errorf("%s: expected false", msg)
	}
}

// AssertContains checks if s contains substr.
func (h *Helper) AssertContains(s, substr, msg string) {
	h.T.Helper()
	if substr == "" {
		return
	}
	if len(substr) > 0 && len(s) >= len(substr) {
		for i := 0; i <= len(s)-len(substr); i++ {
			if s[i:i+len(substr)] == substr {
				return
			}
		}
	}
	h.T.Errorf("%s: %q does not contain %q", msg, s, substr)
}

// =============================================================================
// Parameterized Test Support
// =============================================================================

// TestCase represents a generic test case.
type TestCase struct {
	Name    string
	Setup   func()
	Run     func(t *testing.T)
	Cleanup func()
}

// RunTestCases runs a slice of test cases as subtests.
func RunTestCases(t *testing.T, cases []TestCase) {
	for _, tc := range cases {
		t.Run(tc.Name, func(t *testing.T) {
			if tc.Setup != nil {
				tc.Setup()
			}
			if tc.Cleanup != nil {
				defer tc.Cleanup()
			}
			tc.Run(t)
		})
	}
}

// =============================================================================
// Value Test Cases - Common test data
// =============================================================================

// StringTestCase represents a test case with string input/output.
type StringTestCase struct {
	Name   string
	Input  string
	Want   string
	WantOK bool
}

// CommonStringCases provides standard string test data.
var CommonStringCases = []StringTestCase{
	{"empty", "", "", true},
	{"simple", "Hello", "Hello", true},
	{"spaces", "  spaced  ", "  spaced  ", true},
	{"unicode", "日本語テスト", "日本語テスト", true},
	{"emoji", "🎉 Party 🎊", "🎉 Party 🎊", true},
	{"special_chars", "a < b > c & d \"e\"", "a < b > c & d \"e\"", true},
	{"newlines", "line1\nline2\nline3", "line1\nline2\nline3", true},
	{"tabs", "col1\tcol2\tcol3", "col1\tcol2\tcol3", true},
}

// NumericTestCase represents a test case with numeric input/output.
type NumericTestCase struct {
	Name  string
	Input float64
	Want  float64
}

// CommonNumericCases provides standard numeric test data.
var CommonNumericCases = []NumericTestCase{
	{"zero", 0, 0},
	{"positive_int", 42, 42},
	{"negative_int", -100, -100},
	{"float", 3.14159, 3.14159},
	{"large", 1e10, 1e10},
	{"small", 1e-10, 1e-10},
}

// =============================================================================
// Format Combinations - For formatting tests
// =============================================================================

// FormatOptions represents text formatting options.
type FormatOptions struct {
	Bold      bool
	Italic    bool
	Underline bool
	FontSize  float64
	FontName  string
	Color     string
}

// CommonFormatCombinations provides standard format test combinations.
var CommonFormatCombinations = []FormatOptions{
	{}, // No formatting
	{Bold: true},
	{Italic: true},
	{Underline: true},
	{Bold: true, Italic: true},
	{Bold: true, Italic: true, Underline: true},
	{FontSize: 8},
	{FontSize: 12},
	{FontSize: 24},
	{FontSize: 72},
	{FontName: "Arial"},
	{FontName: "Times New Roman"},
	{Color: "FF0000"},
	{Color: "00FF00"},
	{Color: "0000FF"},
	{Bold: true, FontSize: 14, Color: "FF0000"},
}

// String returns a description of the format options.
func (f FormatOptions) String() string {
	parts := []string{}
	if f.Bold {
		parts = append(parts, "bold")
	}
	if f.Italic {
		parts = append(parts, "italic")
	}
	if f.Underline {
		parts = append(parts, "underline")
	}
	if f.FontSize > 0 {
		parts = append(parts, "size")
	}
	if f.FontName != "" {
		parts = append(parts, "font")
	}
	if f.Color != "" {
		parts = append(parts, "color")
	}
	if len(parts) == 0 {
		return "plain"
	}
	result := parts[0]
	for i := 1; i < len(parts); i++ {
		result += "_" + parts[i]
	}
	return result
}

// =============================================================================
// Cell Reference Test Cases
// =============================================================================

// CellRefTestCase represents a cell reference test case.
type CellRefTestCase struct {
	Name    string
	Ref     string
	Row     int
	Col     int
	WantErr bool
}

// CommonCellRefCases provides standard cell reference test data.
var CommonCellRefCases = []CellRefTestCase{
	{"A1", "A1", 1, 1, false},
	{"Z1", "Z1", 1, 26, false},
	{"AA1", "AA1", 1, 27, false},
	{"AZ1", "AZ1", 1, 52, false},
	{"A10", "A10", 10, 1, false},
	{"A100", "A100", 100, 1, false},
	{"XFD1", "XFD1", 1, 16384, false}, // Max Excel column
	{"invalid", "123", 0, 0, true},
	{"empty", "", 0, 0, true},
}

// =============================================================================
// Range Test Cases
// =============================================================================

// RangeTestCase represents a range test case.
type RangeTestCase struct {
	Name     string
	Ref      string
	StartRow int
	StartCol int
	EndRow   int
	EndCol   int
	WantErr  bool
}

// CommonRangeCases provides standard range test data.
var CommonRangeCases = []RangeTestCase{
	{"single_cell", "A1:A1", 1, 1, 1, 1, false},
	{"row", "A1:D1", 1, 1, 1, 4, false},
	{"column", "A1:A10", 1, 1, 10, 1, false},
	{"block", "B2:D4", 2, 2, 4, 4, false},
	{"large", "A1:Z100", 1, 1, 100, 26, false},
}

// =============================================================================
// Document Type Test Support
// =============================================================================

// DocumentType represents a type of Office document.
type DocumentType int

const (
	// DocTypeWord represents Word documents.
	DocTypeWord DocumentType = iota
	// DocTypePowerPoint represents PowerPoint documents.
	DocTypePowerPoint
	// DocTypeExcel represents Excel documents.
	DocTypeExcel
)

// Extension returns the file extension for the document type.
func (dt DocumentType) Extension() string {
	switch dt {
	case DocTypeWord:
		return ".docx"
	case DocTypePowerPoint:
		return ".pptx"
	case DocTypeExcel:
		return ".xlsx"
	default:
		return ""
	}
}

// String returns the name of the document type.
func (dt DocumentType) String() string {
	switch dt {
	case DocTypeWord:
		return "Word"
	case DocTypePowerPoint:
		return "PowerPoint"
	case DocTypeExcel:
		return "Excel"
	default:
		return "Unknown"
	}
}
