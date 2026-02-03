// Package document provides test utilities and fixtures for document tests.
package document

import (
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

// =============================================================================
// Test Fixtures - Reusable document configurations
// =============================================================================

// TestFixture represents a reusable test document configuration.
type TestFixture struct {
	Name        string
	Setup       func(*Document)
	Description string
}

// CommonFixtures provides standard test fixtures for document tests.
var CommonFixtures = []TestFixture{
	{
		Name:        "empty",
		Description: "Empty document with no content",
		Setup:       func(d *Document) {},
	},
	{
		Name:        "single_paragraph",
		Description: "Document with one plain paragraph",
		Setup: func(d *Document) {
			d.AddParagraph().SetText("Hello World")
		},
	},
	{
		Name:        "multiple_paragraphs",
		Description: "Document with multiple paragraphs",
		Setup: func(d *Document) {
			d.AddParagraph().SetText("First paragraph")
			d.AddParagraph().SetText("Second paragraph")
			d.AddParagraph().SetText("Third paragraph")
		},
	},
	{
		Name:        "formatted_text",
		Description: "Document with various text formatting",
		Setup: func(d *Document) {
			p := d.AddParagraph()
			r1 := p.AddRun()
			r1.SetText("Bold ")
			r1.SetBold(true)

			r2 := p.AddRun()
			r2.SetText("Italic ")
			r2.SetItalic(true)

			r3 := p.AddRun()
			r3.SetText("Underline ")
			r3.SetUnderline(true)

			r4 := p.AddRun()
			r4.SetText("Colored")
			r4.SetColor("FF0000")
		},
	},
	{
		Name:        "headings",
		Description: "Document with heading hierarchy",
		Setup: func(d *Document) {
			h1 := d.AddParagraph()
			h1.SetText("Heading 1")
			h1.SetStyle("Heading1")

			h2 := d.AddParagraph()
			h2.SetText("Heading 2")
			h2.SetStyle("Heading2")

			h3 := d.AddParagraph()
			h3.SetText("Heading 3")
			h3.SetStyle("Heading3")

			p := d.AddParagraph()
			p.SetText("Normal paragraph")
		},
	},
	{
		Name:        "simple_table",
		Description: "Document with a 2x2 table",
		Setup: func(d *Document) {
			tbl := d.AddTable(2, 2)
			tbl.Cell(0, 0).SetText("A1")
			tbl.Cell(0, 1).SetText("B1")
			tbl.Cell(1, 0).SetText("A2")
			tbl.Cell(1, 1).SetText("B2")
		},
	},
	{
		Name:        "mixed_content",
		Description: "Document with paragraphs and tables",
		Setup: func(d *Document) {
			title := d.AddParagraph()
			title.SetText("Document Title")
			title.SetStyle("Heading1")

			intro := d.AddParagraph()
			intro.SetText("Introduction paragraph.")

			tbl := d.AddTable(3, 3)
			for r := 0; r < 3; r++ {
				for c := 0; c < 3; c++ {
					tbl.Cell(r, c).SetText("Cell")
				}
			}

			d.AddParagraph().SetText("Conclusion.")
		},
	},
	{
		Name:        "unicode_content",
		Description: "Document with international text",
		Setup: func(d *Document) {
			texts := []string{
				"English",
				"日本語",
				"中文",
				"한국어",
				"العربية",
				"🎉 Emoji 🎊",
			}
			for _, text := range texts {
				d.AddParagraph().SetText(text)
			}
		},
	},
	{
		Name:        "sdt_content_controls",
		Description: "Document with content controls",
		Setup: func(d *Document) {
			p := d.AddParagraph()
			p.SetStyle("Title")
			p.AddRun().SetText("Content Controls Fixture")

			inlinePara := d.AddParagraph()
			inlinePara.AddRun().SetText("Inline: ")
			inline := inlinePara.AddContentControl("InlineTag", "Inline Alias", "Inline Value")
			inline.SetComboBox([]ContentControlListItem{{DisplayText: "Alpha", Value: "A"}})

			block := d.AddBlockContentControl("BlockTag", "Block Alias", "Block Value")
			block.SetDropDownList([]ContentControlListItem{{DisplayText: "Option A", Value: "A"}, {DisplayText: "Option B", Value: "B"}})

			date := d.AddBlockContentControl("DateTag", "Date Alias", "2026-02-02")
			date.SetDateConfig(ContentControlDateConfig{
				Format:   "yyyy-MM-dd",
				Locale:   "en-US",
				Calendar: "gregorian",
			})
		},
	},
}

// =============================================================================
// Test Helpers - Reduce boilerplate in tests
// =============================================================================

// TestHelper provides common test operations.
type TestHelper struct {
	t       *testing.T
	tempDir string
}

// NewTestHelper creates a new test helper.
func NewTestHelper(t *testing.T) *TestHelper {
	t.Helper()
	return &TestHelper{
		t:       t,
		tempDir: t.TempDir(),
	}
}

// CreateDocument creates a new document, calls setup, and returns it.
// The document is NOT automatically closed - caller must defer doc.Close().
func (h *TestHelper) CreateDocument(setup func(*Document)) *Document {
	h.t.Helper()
	doc, err := New()
	if err != nil {
		h.t.Fatalf("New() error = %v", err)
	}
	if setup != nil {
		setup(doc)
	}
	return doc
}

// SaveDocument saves a document to a temp file and returns the path.
func (h *TestHelper) SaveDocument(doc *Document, name string) string {
	h.t.Helper()
	path := filepath.Join(h.tempDir, name)
	if err := doc.SaveAs(path); err != nil {
		h.t.Fatalf("SaveAs(%q) error = %v", name, err)
	}
	return path
}

// OpenDocument opens a document from a path.
func (h *TestHelper) OpenDocument(path string) *Document {
	h.t.Helper()
	doc, err := Open(filepath.Clean(path))
	if err != nil {
		h.t.Fatalf("Open(%q) error = %v", path, err)
	}
	return doc
}

// OpenFixture opens a document fixture from testdata/word.
func (h *TestHelper) OpenFixture(name string) *Document {
	h.t.Helper()
	return h.OpenDocument(fixturePath(name))
}

func fixturePath(name string) string {
	_, filename, _, ok := runtime.Caller(0)
	if !ok {
		return filepath.Join("..", "..", "testdata", "word", name)
	}
	root := filepath.Join(filepath.Dir(filename), "..", "..")
	return filepath.Join(root, "testdata", "word", name)
}

// RoundTrip creates, saves, and reopens a document.
// Returns the reopened document (caller must close).
func (h *TestHelper) RoundTrip(name string, setup func(*Document)) *Document {
	h.t.Helper()
	doc := h.CreateDocument(setup)
	path := h.SaveDocument(doc, name)
	doc.Close()
	return h.OpenDocument(path)
}

// AssertParagraphCount checks the paragraph count.
func (h *TestHelper) AssertParagraphCount(doc *Document, want int) {
	h.t.Helper()
	got := len(doc.Paragraphs())
	if got != want {
		h.t.Errorf("paragraph count = %d, want %d", got, want)
	}
}

// AssertTableCount checks the table count.
func (h *TestHelper) AssertTableCount(doc *Document, want int) {
	h.t.Helper()
	got := len(doc.Tables())
	if got != want {
		h.t.Errorf("table count = %d, want %d", got, want)
	}
}

// AssertParagraphText checks paragraph text at index.
func (h *TestHelper) AssertParagraphText(doc *Document, index int, want string) {
	h.t.Helper()
	paras := doc.Paragraphs()
	if index >= len(paras) {
		h.t.Fatalf("paragraph index %d out of range (have %d)", index, len(paras))
	}
	got := paras[index].Text()
	if got != want {
		h.t.Errorf("paragraph[%d].Text() = %q, want %q", index, got, want)
	}
}

// =============================================================================
// Validation Helper
// =============================================================================

// Validator provides OOXML validation using the .NET SDK.
type Validator struct {
	path      string
	available bool
}

// NewValidator creates a validator instance.
func NewValidator() *Validator {
	validatorPath := "../../tools/validator/OoxmlValidator/bin/Release/net10.0/OoxmlValidator.dll"
	_, err := os.Stat(validatorPath)
	return &Validator{
		path:      validatorPath,
		available: err == nil,
	}
}

// Available returns true if the validator is available.
func (v *Validator) Available() bool {
	return v.available
}

// Skip skips the test if validator is not available.
func (v *Validator) Skip(t *testing.T) {
	t.Helper()
	if !v.available {
		t.Skip("OOXML validator not available; run 'make validate' to build it")
	}
}

// Validate validates a document file.
func (v *Validator) Validate(t *testing.T, path string) (valid bool, errors string) {
	t.Helper()
	if !v.available {
		t.Skip("OOXML validator not available")
	}

	dotnetRoot := os.Getenv("DOTNET_ROOT")
	if dotnetRoot == "" {
		dotnetRoot = "/home/linuxbrew/.linuxbrew/opt/dotnet/libexec"
	}

	cmd := exec.Command("dotnet", v.path, path, "--json")
	cmd.Env = append(os.Environ(), "DOTNET_ROOT="+dotnetRoot)

	output, err := cmd.CombinedOutput()
	if err != nil {
		if exitErr, ok := err.(*exec.ExitError); ok && exitErr.ExitCode() == 2 {
			return false, string(output)
		}
		return false, string(output)
	}

	return strings.Contains(string(output), `"valid": true`), string(output)
}

// AssertValid validates a document and fails if invalid.
func (v *Validator) AssertValid(t *testing.T, path string) {
	t.Helper()
	valid, output := v.Validate(t, path)
	if !valid {
		t.Errorf("document validation failed:\n%s", output)
	}
}

// =============================================================================
// Parameterized Test Helpers
// =============================================================================

// TextTestCase represents a test case for text content.
type TextTestCase struct {
	Name string
	Text string
}

// CommonTextCases provides standard text test cases.
var CommonTextCases = []TextTestCase{
	{"empty", ""},
	{"simple", "Hello World"},
	{"whitespace", "  spaces  "},
	{"tabs", "col1\tcol2\tcol3"},
	{"unicode_japanese", "日本語テキスト"},
	{"unicode_chinese", "中文文本"},
	{"unicode_arabic", "النص العربي"},
	{"emoji", "🎉 Party 🎊"},
	{"special_chars", "a < b > c & d"},
	{"long_text", strings.Repeat("Lorem ipsum dolor sit amet. ", 100)},
}

// HeadingStyleTestCase represents a test case for paragraph heading styles.
type HeadingStyleTestCase struct {
	Name         string
	StyleID      string
	IsHeading    bool
	HeadingLevel int
}

// CommonHeadingStyleCases provides standard heading style test cases.
var CommonHeadingStyleCases = []HeadingStyleTestCase{
	{"heading1", "Heading1", true, 1},
	{"heading2", "Heading2", true, 2},
	{"heading3", "Heading3", true, 3},
	{"heading4", "Heading4", true, 4},
	{"heading5", "Heading5", true, 5},
	{"heading6", "Heading6", true, 6},
	{"heading7", "Heading7", true, 7},
	{"heading8", "Heading8", true, 8},
	{"heading9", "Heading9", true, 9},
	{"normal", "Normal", false, 0},
	{"title", "Title", false, 0},
	{"subtitle", "Subtitle", false, 0},
	{"empty", "", false, 0},
}

// FontSizeTestCase represents a test case for font sizes.
type FontSizeTestCase struct {
	Points    float64
	HalfPts   int64 // Expected internal representation
}

// CommonFontSizeCases provides standard font size test cases.
var CommonFontSizeCases = []FontSizeTestCase{
	{8, 16},
	{9, 18},
	{10, 20},
	{10.5, 21},
	{11, 22},
	{12, 24},
	{14, 28},
	{16, 32},
	{18, 36},
	{20, 40},
	{24, 48},
	{28, 56},
	{36, 72},
	{48, 96},
	{72, 144},
}

// AlignmentTestCase represents a test case for paragraph alignment.
type AlignmentTestCase struct {
	Name  string
	Value string
}

// CommonAlignmentCases provides standard alignment test cases.
var CommonAlignmentCases = []AlignmentTestCase{
	{"left", "left"},
	{"center", "center"},
	{"right", "right"},
	{"justify", "both"},
}

// ColorTestCase represents a test case for colors.
type ColorTestCase struct {
	Name  string
	Input string
	Want  string // Expected normalized value
}

// CommonColorCases provides standard color test cases.
var CommonColorCases = []ColorTestCase{
	{"red", "FF0000", "FF0000"},
	{"green", "00FF00", "00FF00"},
	{"blue", "0000FF", "0000FF"},
	{"black", "000000", "000000"},
	{"white", "FFFFFF", "FFFFFF"},
	{"with_hash", "#FF0000", "FF0000"},
	{"lowercase", "ff0000", "ff0000"},
}

// TableDimensionTestCase represents a test case for table dimensions.
type TableDimensionTestCase struct {
	Name string
	Rows int
	Cols int
}

// CommonTableDimensionCases provides standard table dimension test cases.
var CommonTableDimensionCases = []TableDimensionTestCase{
	{"1x1", 1, 1},
	{"1x5", 1, 5},
	{"5x1", 5, 1},
	{"2x2", 2, 2},
	{"3x3", 3, 3},
	{"5x5", 5, 5},
	{"10x3", 10, 3},
	{"3x10", 3, 10},
}

// =============================================================================
// Benchmark Helpers
// =============================================================================

// BenchmarkHelper provides utilities for benchmark tests.
type BenchmarkHelper struct {
	b *testing.B
}

// NewBenchmarkHelper creates a new benchmark helper.
func NewBenchmarkHelper(b *testing.B) *BenchmarkHelper {
	return &BenchmarkHelper{b: b}
}

// ResetTimer resets the benchmark timer after setup.
func (h *BenchmarkHelper) ResetTimer() {
	h.b.ResetTimer()
}

// StopTimer stops the timer during teardown.
func (h *BenchmarkHelper) StopTimer() {
	h.b.StopTimer()
}
