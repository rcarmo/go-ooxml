// Package document provides an enhanced test framework for comprehensive testing.
// This file provides:
// - Extended fixtures for all Phase 3 features
// - Parameterized test case builders
// - Assertion helpers with better error messages
// - Test matrix generators for combinatorial testing
package document

import (
	"fmt"
	"path/filepath"
	"testing"
)

// =============================================================================
// Phase 3 Fixtures - Track Changes, Comments, Styles, Headers/Footers
// =============================================================================

// Phase3Fixtures provides fixtures specifically for Phase 3 features.
var Phase3Fixtures = []TestFixture{
	// Track Changes Fixtures
	{
		Name:        "track_changes_insert",
		Description: "Document with tracked insertions",
		Setup: func(d Document) {
			d.EnableTrackChanges("Editor")
			p := d.AddParagraph()
			p.InsertTrackedText("This text was inserted")
		},
	},
	{
		Name:        "track_changes_delete",
		Description: "Document with tracked deletions",
		Setup: func(d Document) {
			p := d.AddParagraph()
			p.SetText("Text to be deleted")
			d.EnableTrackChanges("Editor")
			p.DeleteTrackedText(0)
		},
	},
	{
		Name:        "track_changes_mixed",
		Description: "Document with mixed tracked changes",
		Setup: func(d Document) {
			d.EnableTrackChanges("Editor1")
			p := d.AddParagraph()
			p.InsertTrackedText("First insertion")
			d.SetTrackAuthor("Editor2")
			p.InsertTrackedText("Second insertion")
		},
	},
	{
		Name:        "track_changes_multiple_authors",
		Description: "Document with changes from multiple authors",
		Setup: func(d Document) {
			d.EnableTrackChanges("Author A")
			d.AddParagraph().InsertTrackedText("From Author A")
			d.SetTrackAuthor("Author B")
			d.AddParagraph().InsertTrackedText("From Author B")
			d.SetTrackAuthor("Author C")
			d.AddParagraph().InsertTrackedText("From Author C")
		},
	},

	// Comments Fixtures
	{
		Name:        "single_comment",
		Description: "Document with one comment",
		Setup: func(d Document) {
			d.AddParagraph().SetText("Commented text")
			_, _ = d.Comments().Add("Please check this section", "Reviewer", "")
		},
	},
	{
		Name:        "multiple_comments",
		Description: "Document with multiple comments",
		Setup: func(d Document) {
			d.AddParagraph().SetText("First paragraph")
			_, _ = d.Comments().Add("First comment", "Reviewer 1", "")
			d.AddParagraph().SetText("Second paragraph")
			_, _ = d.Comments().Add("Second comment", "Reviewer 2", "")
			_, _ = d.Comments().Add("Third comment", "Reviewer 1", "")
		},
	},
	{
		Name:        "comment_with_initials",
		Description: "Document with comment including initials",
		Setup: func(d Document) {
			d.AddParagraph().SetText("Text")
			c, _ := d.Comments().Add("A comment", "John Doe", "")
			c.SetInitials("JD")
		},
	},

	// Styles Fixtures
	{
		Name:        "custom_paragraph_style",
		Description: "Document with custom paragraph style",
		Setup: func(d Document) {
			s := d.AddParagraphStyle("CustomStyle", "Custom Style")
			s.SetBold(true)
			s.SetFontSize(14)
			s.SetColor("0000FF")
			
			p := d.AddParagraph()
			p.SetStyle("CustomStyle")
			p.SetText("Styled text")
		},
	},
	{
		Name:        "style_inheritance",
		Description: "Document with inherited styles",
		Setup: func(d Document) {
			base := d.AddParagraphStyle("BaseStyle", "Base Style")
			base.SetFontSize(12)
			base.SetFontName("Arial")
			
			derived := d.AddParagraphStyle("DerivedStyle", "Derived Style")
			derived.SetBasedOn("BaseStyle")
			derived.SetBold(true)
			
			d.AddParagraph().SetStyle("DerivedStyle")
		},
	},
	{
		Name:        "multiple_style_types",
		Description: "Document with paragraph, character, and table styles",
		Setup: func(d Document) {
			d.AddParagraphStyle("ParaStyle", "Paragraph Style")
			d.AddCharacterStyle("CharStyle", "Character Style")
			d.AddTableStyle("TableStyle", "Table Style")
		},
	},

	// Headers/Footers Fixtures
	{
		Name:        "default_header_footer",
		Description: "Document with default header and footer",
		Setup: func(d Document) {
			h := d.AddHeader(HeaderFooterDefault)
			h.SetText("Default Header")
			f := d.AddFooter(HeaderFooterDefault)
			f.SetText("Default Footer")
		},
	},
	{
		Name:        "first_page_header_footer",
		Description: "Document with first page header/footer",
		Setup: func(d Document) {
			d.AddHeader(HeaderFooterDefault).SetText("Default Header")
			d.AddHeader(HeaderFooterFirst).SetText("First Page Header")
			d.AddFooter(HeaderFooterDefault).SetText("Default Footer")
			d.AddFooter(HeaderFooterFirst).SetText("First Page Footer")
		},
	},
	{
		Name:        "all_header_footer_types",
		Description: "Document with all header/footer types",
		Setup: func(d Document) {
			d.AddHeader(HeaderFooterDefault).SetText("Default")
			d.AddHeader(HeaderFooterFirst).SetText("First")
			d.AddHeader(HeaderFooterEven).SetText("Even")
			d.AddFooter(HeaderFooterDefault).SetText("Default")
			d.AddFooter(HeaderFooterFirst).SetText("First")
			d.AddFooter(HeaderFooterEven).SetText("Even")
		},
	},

	// Combined Features Fixtures
	{
		Name:        "full_featured_document",
		Description: "Document using all Phase 3 features",
		Setup: func(d Document) {
			// Add styles
			d.AddParagraphStyle("Title", "Title").SetFontSize(24)
			d.AddParagraphStyle("Body", "Body Text").SetFontSize(11)
			
			// Add header/footer
			d.AddHeader(HeaderFooterDefault).SetText("Document Title")
			d.AddFooter(HeaderFooterDefault).SetText("Page Footer")
			
			// Add content with styles
			title := d.AddParagraph()
			title.SetStyle("Title")
			title.SetText("Document Title")
			
			// Add tracked changes
			d.EnableTrackChanges("Editor")
			body := d.AddParagraph()
			body.SetStyle("Body")
			body.InsertTrackedText("This is tracked content")
			
			// Add comment
			_, _ = d.Comments().Add("Please review this document", "Reviewer", "")
		},
	},
}

// =============================================================================
// Parameterized Test Case Builders
// =============================================================================

// TrackChangesTestCase represents a parameterized track changes test.
type TrackChangesTestCase struct {
	Name           string
	Author         string
	Text           string
	ExpectEnabled  bool
	ExpectRevision bool
}

// TrackChangesTestCases returns standard test cases for track changes.
func TrackChangesTestCases() []TrackChangesTestCase {
	return []TrackChangesTestCase{
		{"basic insert", "Editor", "inserted text", true, true},
		{"empty author", "", "text", true, true},
		{"unicode author", "日本語", "text", true, true},
		{"long text", "Author", "This is a very long piece of text that spans multiple words and sentences to test handling of longer content.", true, true},
		{"special chars in text", "Editor", "<>&\"'", true, true},
	}
}

// CommentTestCase represents a parameterized comment test.
type CommentTestCase struct {
	Name     string
	Author   string
	Initials string
	Text     string
}

// CommentTestCases returns standard test cases for comments.
func CommentTestCases() []CommentTestCase {
	return []CommentTestCase{
		{"basic comment", "John Doe", "JD", "This is a comment"},
		{"empty initials", "Jane Smith", "", "Another comment"},
		{"unicode author", "田中太郎", "田", "Japanese comment"},
		{"long comment", "Author", "A", "This is a very long comment that contains multiple sentences. It should be handled correctly regardless of its length."},
		{"special chars", "Author <>&", "A", "Comment with <special> & \"chars\""},
	}
}

// StyleTestCase represents a parameterized style test.
type StyleTestCase struct {
	Name      string
	StyleID   string
	StyleName string
	StyleType StyleType
	Bold      bool
	Italic    bool
	FontSize  float64
	Color     string
}

// StyleTestCases returns standard test cases for styles.
func StyleTestCases() []StyleTestCase {
	return []StyleTestCase{
		{"basic paragraph", "Para1", "Paragraph One", StyleTypeParagraph, false, false, 11, ""},
		{"bold style", "BoldStyle", "Bold Text", StyleTypeParagraph, true, false, 11, ""},
		{"italic style", "ItalicStyle", "Italic Text", StyleTypeParagraph, false, true, 11, ""},
		{"colored style", "Colored", "Colored Text", StyleTypeParagraph, false, false, 11, "FF0000"},
		{"large font", "Large", "Large Text", StyleTypeParagraph, false, false, 24, ""},
		{"character style", "CharBold", "Bold Characters", StyleTypeCharacter, true, false, 0, ""},
		{"table style", "TableGrid", "Table Grid", StyleTypeTable, false, false, 0, ""},
		{"numbering style", "Numbered", "Numbered List", StyleTypeNumbering, false, false, 0, ""},
	}
}

// HeaderFooterTestCase represents a parameterized header/footer test.
type HeaderFooterTestCase struct {
	Name   string
	Type   HeaderFooterType
	Text   string
	IsHeader bool
}

// HeaderFooterTestCases returns standard test cases for headers/footers.
func HeaderFooterTestCases() []HeaderFooterTestCase {
	return []HeaderFooterTestCase{
		{"default header", HeaderFooterDefault, "Default Header", true},
		{"first header", HeaderFooterFirst, "First Page Header", true},
		{"even header", HeaderFooterEven, "Even Page Header", true},
		{"default footer", HeaderFooterDefault, "Default Footer", false},
		{"first footer", HeaderFooterFirst, "First Page Footer", false},
		{"even footer", HeaderFooterEven, "Even Page Footer", false},
	}
}

// =============================================================================
// Enhanced Test Helper Methods
// =============================================================================

// WithTrackChanges runs a test function with track changes enabled.
func (h *TestHelper) WithTrackChanges(author string, fn func(Document)) Document {
	h.t.Helper()
	doc := h.CreateDocument(func(d Document) {
		d.EnableTrackChanges(author)
		fn(d)
	})
	return doc
}

// WithComments runs a test function and adds specified comments.
func (h *TestHelper) WithComments(comments []CommentTestCase, fn func(Document)) Document {
	h.t.Helper()
	doc := h.CreateDocument(func(d Document) {
		fn(d)
		for _, c := range comments {
			comment, _ := d.Comments().Add(c.Text, c.Author, "")
			if c.Initials != "" {
				comment.SetInitials(c.Initials)
			}
		}
	})
	return doc
}

// RoundTripWithFixture creates a document using a fixture, saves, and reopens it.
func (h *TestHelper) RoundTripWithFixture(fixture TestFixture) Document {
	h.t.Helper()
	return h.RoundTrip(fixture.Name+".docx", fixture.Setup)
}

// AssertTrackChangesEnabled checks track changes state.
func (h *TestHelper) AssertTrackChangesEnabled(doc Document, want bool) {
	h.t.Helper()
	if got := doc.TrackChangesEnabled(); got != want {
		h.t.Errorf("TrackChangesEnabled() = %v, want %v", got, want)
	}
}

// AssertTrackAuthor checks the track changes author.
func (h *TestHelper) AssertTrackAuthor(doc Document, want string) {
	h.t.Helper()
	if got := doc.TrackAuthor(); got != want {
		h.t.Errorf("TrackAuthor() = %q, want %q", got, want)
	}
}

// AssertRevisionCount checks the number of revisions.
func (h *TestHelper) AssertRevisionCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.AllRevisions())
	if got != want {
		h.t.Errorf("len(AllRevisions()) = %d, want %d", got, want)
	}
}

// AssertInsertionCount checks the number of insertions.
func (h *TestHelper) AssertInsertionCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.Insertions())
	if got != want {
		h.t.Errorf("len(Insertions()) = %d, want %d", got, want)
	}
}

// AssertDeletionCount checks the number of deletions.
func (h *TestHelper) AssertDeletionCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.Deletions())
	if got != want {
		h.t.Errorf("len(Deletions()) = %d, want %d", got, want)
	}
}

// AssertCommentCount checks the number of comments.
func (h *TestHelper) AssertCommentCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.Comments().All())
	if got != want {
		h.t.Errorf("len(Comments()) = %d, want %d", got, want)
	}
}

// AssertStyleCount checks the number of styles.
func (h *TestHelper) AssertStyleCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.Styles().List())
	if got != want {
		h.t.Errorf("len(Styles()) = %d, want %d", got, want)
	}
}

// AssertHeaderCount checks the number of headers.
func (h *TestHelper) AssertHeaderCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.Headers())
	if got != want {
		h.t.Errorf("len(Headers()) = %d, want %d", got, want)
	}
}

// AssertFooterCount checks the number of footers.
func (h *TestHelper) AssertFooterCount(doc Document, want int) {
	h.t.Helper()
	got := len(doc.Footers())
	if got != want {
		h.t.Errorf("len(Footers()) = %d, want %d", got, want)
	}
}

// =============================================================================
// Test Matrix Generators
// =============================================================================

// RunFixtureTests runs a test function against all fixtures.
func RunFixtureTests(t *testing.T, fixtures []TestFixture, testFn func(*testing.T, *TestHelper, TestFixture)) {
	for _, fixture := range fixtures {
		t.Run(fixture.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			testFn(t, h, fixture)
		})
	}
}

// RunTrackChangesTests runs parameterized track changes tests.
func RunTrackChangesTests(t *testing.T, cases []TrackChangesTestCase, testFn func(*testing.T, *TestHelper, TrackChangesTestCase)) {
	for _, tc := range cases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			testFn(t, h, tc)
		})
	}
}

// RunCommentTests runs parameterized comment tests.
func RunCommentTests(t *testing.T, cases []CommentTestCase, testFn func(*testing.T, *TestHelper, CommentTestCase)) {
	for _, tc := range cases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			testFn(t, h, tc)
		})
	}
}

// RunStyleTests runs parameterized style tests.
func RunStyleTests(t *testing.T, cases []StyleTestCase, testFn func(*testing.T, *TestHelper, StyleTestCase)) {
	for _, tc := range cases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			testFn(t, h, tc)
		})
	}
}

// RunHeaderFooterTests runs parameterized header/footer tests.
func RunHeaderFooterTests(t *testing.T, cases []HeaderFooterTestCase, testFn func(*testing.T, *TestHelper, HeaderFooterTestCase)) {
	for _, tc := range cases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			testFn(t, h, tc)
		})
	}
}

// =============================================================================
// Combinatorial Test Generators
// =============================================================================

// FormatCombination represents a combination of formatting options.
type FormatCombination struct {
	Bold      bool
	Italic    bool
	Underline bool
	FontSize  float64
	Color     string
}

// AllFormatCombinations generates all combinations of formatting.
func AllFormatCombinations() []FormatCombination {
	var combinations []FormatCombination
	bools := []bool{false, true}
	sizes := []float64{11, 14, 18}
	colors := []string{"", "FF0000", "0000FF"}
	
	for _, bold := range bools {
		for _, italic := range bools {
			for _, underline := range bools {
				for _, size := range sizes {
					for _, color := range colors {
						combinations = append(combinations, FormatCombination{
							Bold:      bold,
							Italic:    italic,
							Underline: underline,
							FontSize:  size,
							Color:     color,
						})
					}
				}
			}
		}
	}
	return combinations
}

// FormatCombinationName returns a descriptive name for a format combination.
func (fc FormatCombination) String() string {
	parts := []string{}
	if fc.Bold {
		parts = append(parts, "bold")
	}
	if fc.Italic {
		parts = append(parts, "italic")
	}
	if fc.Underline {
		parts = append(parts, "underline")
	}
	parts = append(parts, fmt.Sprintf("%.0fpt", fc.FontSize))
	if fc.Color != "" {
		parts = append(parts, fc.Color)
	}
	if len(parts) == 1 {
		return "plain_" + parts[0]
	}
	return filepath.Join(parts...)
}

// ApplyToRun applies the formatting to a run.
func (fc FormatCombination) ApplyToRun(r Run) {
	r.SetBold(fc.Bold)
	r.SetItalic(fc.Italic)
	r.SetUnderline(fc.Underline)
	if fc.FontSize > 0 {
		r.SetFontSize(fc.FontSize)
	}
	if fc.Color != "" {
		r.SetColor(fc.Color)
	}
}
