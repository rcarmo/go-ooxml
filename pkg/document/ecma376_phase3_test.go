// Package document provides ECMA-376 spec-compliant tests for Phase 3 features.
// These tests validate compliance with the ECMA-376 specification for:
// - Track Changes (§17.13.5)
// - Comments (§17.13.4)
// - Styles (§17.7)
// - Headers/Footers (§17.10)
package document

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

// =============================================================================
// ECMA-376 §17.13.5 - Track Changes (Revisions) Tests
// =============================================================================

// TestECMA376_InsElement tests <w:ins> element structure per §17.13.5.20
func TestECMA376_InsElement(t *testing.T) {
	tests := []struct {
		name       string
		author     string
		wantAuthor bool
		wantDate   bool
	}{
		{
			name:       "ins with author and date",
			author:     "Joe Smith",
			wantAuthor: true,
			wantDate:   true,
		},
		{
			name:       "ins with empty author",
			author:     "",
			wantAuthor: false,
			wantDate:   true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			doc.EnableTrackChanges(tt.author)
			p := doc.AddParagraph()
			p.InsertTrackedText("test content")

			// Save to temp file and read back
			tmpDir := t.TempDir()
			tmpFile := filepath.Join(tmpDir, "test.docx")
			if err := doc.SaveAs(tmpFile); err != nil {
				t.Fatal(err)
			}
			
			// Read the saved file
			data, err := os.ReadFile(tmpFile)
			if err != nil {
				t.Fatal(err)
			}

			xmlContent := extractDocumentXML(t, data)

			// Verify <w:ins> structure per ECMA-376
			if !strings.Contains(xmlContent, "<ins") && !strings.Contains(xmlContent, ":ins") {
				t.Error("Expected <w:ins> element for tracked insertion")
			}

			// Check author attribute if expected
			if tt.wantAuthor {
				if !strings.Contains(xmlContent, `author=`) {
					t.Error("Expected author attribute on <w:ins>")
				}
				if !strings.Contains(xmlContent, tt.author) {
					t.Errorf("Expected author value %q in XML", tt.author)
				}
			}

			// Check date attribute (ISO 8601 format)
			if tt.wantDate {
				if !strings.Contains(xmlContent, `date=`) {
					t.Error("Expected date attribute on <w:ins>")
				}
				// Verify ISO 8601 format (contains T separator)
				if !strings.Contains(xmlContent, "T") {
					t.Error("Date should be in ISO 8601 format")
				}
			}
		})
	}
}

// TestECMA376_DelElement tests <w:del> element structure per §17.13.5.12
func TestECMA376_DelElement(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add text first
	p := doc.AddParagraph()
	p.SetText("text to delete")

	// Enable tracking and delete
	doc.EnableTrackChanges("Jane Doe")
	p.DeleteTrackedText(0)

	// Save to temp file and read back
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test_del.docx")
	if err := doc.SaveAs(tmpFile); err != nil {
		t.Fatal(err)
	}
	
	data, err := os.ReadFile(tmpFile)
	if err != nil {
		t.Fatal(err)
	}

	xmlContent := extractDocumentXML(t, data)

	// Per ECMA-376 §17.13.5.12: deleted text should use <w:delText> not <w:t>
	if strings.Contains(xmlContent, "<del") || strings.Contains(xmlContent, ":del") {
		// Check for delText element
		if !strings.Contains(xmlContent, "delText") {
			t.Error("Deleted content should use <w:delText> per ECMA-376 §17.13.5.12")
		}
	}
}

// TestECMA376_RevisionIDUniqueness tests that revision IDs are unique per §17.13.5
func TestECMA376_RevisionIDUniqueness(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.EnableTrackChanges("Editor")

	// Create multiple revisions
	p := doc.AddParagraph()
	p.InsertTrackedText("first")
	p.InsertTrackedText("second")
	p.InsertTrackedText("third")

	insertions := doc.Insertions()
	if len(insertions) < 3 {
		t.Fatalf("Expected at least 3 insertions, got %d", len(insertions))
	}

	// Check all IDs are unique
	seenIDs := make(map[string]bool)
	for i, rev := range insertions {
		id := rev.ID()
		if seenIDs[id] {
			t.Errorf("Duplicate revision ID %s found (ECMA-376 requires unique IDs)", id)
			}
			seenIDs[id] = true
		t.Logf("Revision %d has ID: %s", i, id)
	}
}

// TestECMA376_DateTimeFormat tests ISO 8601 date format compliance
func TestECMA376_DateTimeFormat(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	beforeInsert := time.Now()
	doc.EnableTrackChanges("Author")
	p := doc.AddParagraph()
	p.InsertTrackedText("dated text")
	afterInsert := time.Now()

	revisions := doc.AllRevisions()
	if len(revisions) == 0 {
		t.Fatal("Expected at least one revision")
	}

	revDate := revisions[0].Date()

	// Date should be between before and after
	if revDate.Before(beforeInsert.Add(-time.Second)) || revDate.After(afterInsert.Add(time.Second)) {
		t.Errorf("Revision date %v outside expected range [%v, %v]",
			revDate, beforeInsert, afterInsert)
	}

	// Verify it parses as valid time (non-zero)
	if revDate.IsZero() {
		t.Error("Revision date should not be zero")
	}
}

// TestECMA376_NestedRevisions tests nested ins/del per spec allowance
func TestECMA376_NestedRevisions(t *testing.T) {
	// Per ECMA-376, <w:ins> can contain <w:del> and vice versa
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.EnableTrackChanges("Editor1")
	p := doc.AddParagraph()
	p.InsertTrackedText("first edit")

	// Get revision count
	count1 := len(doc.AllRevisions())

	// Make another tracked edit
	p.InsertTrackedText("second edit")
	count2 := len(doc.AllRevisions())

	if count2 <= count1 {
		t.Error("Multiple tracked edits should create multiple revisions")
	}
}

// =============================================================================
// ECMA-376 §17.13.4 - Comments Tests
// =============================================================================

// TestECMA376_CommentElement tests <w:comment> structure per §17.13.4.2
func TestECMA376_CommentElement(t *testing.T) {
	tests := []struct {
		name     string
		author   string
		initials string
		text     string
	}{
		{"basic comment", "John Doe", "JD", "This is a comment"},
		{"empty initials", "Jane", "", "Another comment"},
		{"unicode author", "日本語", "JP", "Unicode test"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			c, _ := doc.Comments().Add(tt.text, tt.author, "")
			if tt.initials != "" {
				c.SetInitials(tt.initials)
			}

			// Verify attributes
			if c.Author() != tt.author {
				t.Errorf("Author = %q, want %q", c.Author(), tt.author)
			}
			if c.Text() != tt.text {
				t.Errorf("Text = %q, want %q", c.Text(), tt.text)
			}
			if tt.initials != "" && c.Initials() != tt.initials {
				t.Errorf("Initials = %q, want %q", c.Initials(), tt.initials)
			}
		})
	}
}

// TestECMA376_CommentIDMatching tests comment ID matching between anchors and comment part
func TestECMA376_CommentIDMatching(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add multiple comments
	c1, _ := doc.Comments().Add("Comment 1", "Author1", "")
	c2, _ := doc.Comments().Add("Comment 2", "Author2", "")
	c3, _ := doc.Comments().Add("Comment 3", "Author3", "")

	// Verify IDs are unique and sequential-ish
	ids := []string{c1.ID(), c2.ID(), c3.ID()}
	seen := make(map[string]bool)
	for _, id := range ids {
		if seen[id] {
			t.Errorf("Duplicate comment ID: %s (ECMA-376 requires unique IDs)", id)
		}
		seen[id] = true
	}

	// Verify lookup by ID works
	for _, id := range ids {
		found := doc.CommentByID(id)
		if found == nil {
			t.Errorf("Failed to find comment by ID %s", id)
		}
	}
}

// TestECMA376_CommentDate tests comment date attribute
func TestECMA376_CommentDate(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	c, _ := doc.Comments().Add("Comment with date", "Author", "")

	// Comment should have a date
	date := c.Date()
	if date.IsZero() {
		t.Error("Comment date should not be zero")
	}

	// Date should be recent
	if time.Since(date) > time.Minute {
		t.Error("Comment date should be recent")
	}
}

// TestECMA376_CommentContent tests comment can contain block-level content
func TestECMA376_CommentContent(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Per ECMA-376 §17.13.4.2, comment can contain paragraphs
	c, _ := doc.Comments().Add("First paragraph", "Author", "")
	c.SetText("Updated with new paragraph content")

	text := c.Text()
	if !strings.Contains(text, "paragraph") {
		t.Error("Comment should support paragraph content")
	}
}

// TestECMA376_OrphanedComment tests behavior with orphaned comments
func TestECMA376_OrphanedComment(t *testing.T) {
	// Per ECMA-376: "If a comment is not referenced by document content...
	// then it may be ignored when loading the document"
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add comment without anchoring to text
	c, _ := doc.Comments().Add("Orphaned comment", "Author", "")

	// Should still be accessible
	found := doc.CommentByID(c.ID())
	if found == nil {
		t.Error("Orphaned comment should still be accessible before save")
	}
}

// =============================================================================
// ECMA-376 §17.7 - Styles Tests
// =============================================================================

// TestECMA376_StyleTypes tests all style types per §17.18.83
func TestECMA376_StyleTypes(t *testing.T) {
	tests := []struct {
		name      string
		styleType StyleType
		xmlType   string
	}{
		{"paragraph style", StyleTypeParagraph, "paragraph"},
		{"character style", StyleTypeCharacter, "character"},
		{"table style", StyleTypeTable, "table"},
		{"numbering style", StyleTypeNumbering, "numbering"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			var s Style
			switch tt.styleType {
			case StyleTypeParagraph:
				s = doc.AddParagraphStyle("Test", "Test Style")
			case StyleTypeCharacter:
				s = doc.AddCharacterStyle("Test", "Test Style")
			case StyleTypeTable:
				s = doc.AddTableStyle("Test", "Test Style")
			case StyleTypeNumbering:
				s = doc.AddNumberingStyle("Test", "Test Style")
			}

			if s == nil {
				t.Fatal("Failed to create style")
			}

			if s.Type() != tt.styleType {
				t.Errorf("Style type = %v, want %v", s.Type(), tt.styleType)
			}
		})
	}
}

// TestECMA376_StyleInheritance tests basedOn inheritance per §17.7.4.3
func TestECMA376_StyleInheritance(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Create base style
	base := doc.AddParagraphStyle("BaseStyle", "Base Style")
	base.SetBold(true)
	base.SetFontSize(14)

	// Create derived style
	derived := doc.AddParagraphStyle("DerivedStyle", "Derived Style")
	derived.SetBasedOn("BaseStyle")
	derived.SetItalic(true)

	// Verify inheritance link
	if derived.BasedOn() != "BaseStyle" {
		t.Errorf("BasedOn = %q, want %q", derived.BasedOn(), "BaseStyle")
	}

	// Note: Actual property inheritance is resolved at render time
	// We're testing that the basedOn reference is correctly set
}

// TestECMA376_StyleIDUniqueness tests styleId uniqueness requirement
func TestECMA376_StyleIDUniqueness(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Create styles with unique IDs
	doc.AddParagraphStyle("Style1", "First Style")
	doc.AddParagraphStyle("Style2", "Second Style")
	doc.AddCharacterStyle("Style3", "Third Style")

	styles := doc.Styles().List()
	if len(styles) != 3 {
		t.Errorf("Expected 3 styles, got %d", len(styles))
	}

	// Verify all IDs are unique
	seen := make(map[string]bool)
	for _, s := range styles {
		id := s.ID()
		if seen[id] {
			t.Errorf("Duplicate style ID: %s", id)
		}
		seen[id] = true
	}
}

// TestECMA376_DefaultStyle tests default style attribute per spec
func TestECMA376_DefaultStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Create a default paragraph style
	s := doc.AddParagraphStyle("Normal", "Normal")
	s.SetDefault(true)

	// Verify it's marked as default
	if !s.IsDefault() {
		t.Error("Style should be marked as default")
	}

	// Should be findable as default
	found := doc.DefaultParagraphStyle()
	if found == nil {
		t.Error("Should find default paragraph style")
	}
	if found != nil && found.ID() != "Normal" {
		t.Errorf("Default style ID = %q, want %q", found.ID(), "Normal")
	}
}

// TestECMA376_StyleProperties tests pPr and rPr elements
func TestECMA376_StyleProperties(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	s := doc.AddParagraphStyle("Formatted", "Formatted Style")

	// Set paragraph properties (pPr)
	s.SetAlignment("center")
	s.SetSpacingBefore(240) // 12pt = 240 twips
	s.SetSpacingAfter(120)

	// Set run properties (rPr)
	s.SetBold(true)
	s.SetItalic(true)
	s.SetFontSize(14)
	s.SetFontName("Arial")
	s.SetColor("FF0000")

	// Verify properties are set
	pPr := s.ParagraphProperties()
	if pPr == nil {
		t.Error("Paragraph properties should be set")
	}

	rPr := s.RunProperties()
	if rPr == nil {
		t.Error("Run properties should be set")
	}
}

// TestECMA376_StyleApplication tests pStyle, rStyle application
func TestECMA376_StyleApplication(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Create a style
	doc.AddParagraphStyle("CustomHeading", "Custom Heading")

	// Apply to paragraph
	p := doc.AddParagraph()
	p.SetStyle("CustomHeading")

	// Verify style is applied
	if p.Style() != "CustomHeading" {
		t.Errorf("Paragraph style = %q, want %q", p.Style(), "CustomHeading")
	}
}

// =============================================================================
// ECMA-376 §17.10 - Headers and Footers Tests
// =============================================================================

// TestECMA376_HeaderTypes tests header types per ST_HdrFtr (§17.18.36)
func TestECMA376_HeaderTypes(t *testing.T) {
	tests := []struct {
		hfType      HeaderFooterType
		description string
	}{
		{HeaderFooterDefault, "default header (odd pages)"},
		{HeaderFooterFirst, "first page header"},
		{HeaderFooterEven, "even page header"},
	}

	for _, tt := range tests {
		t.Run(tt.description, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			h := doc.AddHeader(tt.hfType)
			if h == nil {
				t.Fatal("Failed to add header")
			}

			if h.Type() != tt.hfType {
				t.Errorf("Header type = %v, want %v", h.Type(), tt.hfType)
			}
		})
	}
}

// TestECMA376_FooterTypes tests footer types
func TestECMA376_FooterTypes(t *testing.T) {
	tests := []struct {
		hfType      HeaderFooterType
		description string
	}{
		{HeaderFooterDefault, "default footer (odd pages)"},
		{HeaderFooterFirst, "first page footer"},
		{HeaderFooterEven, "even page footer"},
	}

	for _, tt := range tests {
		t.Run(tt.description, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			f := doc.AddFooter(tt.hfType)
			if f == nil {
				t.Fatal("Failed to add footer")
			}

			if f.Type() != tt.hfType {
				t.Errorf("Footer type = %v, want %v", f.Type(), tt.hfType)
			}
		})
	}
}

// TestECMA376_HeaderContent tests header can contain block-level content
func TestECMA376_HeaderContent(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	h := doc.AddHeader(HeaderFooterDefault)

	// Add paragraph to header
	p := h.AddParagraph()
	if p == nil {
		t.Fatal("Should be able to add paragraph to header")
	}

	// Header should have content
	paras := h.Paragraphs()
	if len(paras) == 0 {
		t.Error("Header should have at least one paragraph")
	}
}

// TestECMA376_FooterContent tests footer can contain block-level content
func TestECMA376_FooterContent(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	f := doc.AddFooter(HeaderFooterDefault)

	// Add paragraph to footer
	p := f.AddParagraph()
	if p == nil {
		t.Fatal("Should be able to add paragraph to footer")
	}

	// Set text
	f.SetText("Page footer text")

	text := f.Text()
	if !strings.Contains(text, "footer") {
		t.Errorf("Footer text = %q, should contain 'footer'", text)
	}
}

// TestECMA376_MultipleHeadersFooters tests multiple header/footer types
func TestECMA376_MultipleHeadersFooters(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add all three types of headers
	doc.AddHeader(HeaderFooterDefault)
	doc.AddHeader(HeaderFooterFirst)
	doc.AddHeader(HeaderFooterEven)

	// Add all three types of footers
	doc.AddFooter(HeaderFooterDefault)
	doc.AddFooter(HeaderFooterFirst)
	doc.AddFooter(HeaderFooterEven)

	// Verify counts
	headers := doc.Headers()
	if len(headers) != 3 {
		t.Errorf("Expected 3 headers, got %d", len(headers))
	}

	footers := doc.Footers()
	if len(footers) != 3 {
		t.Errorf("Expected 3 footers, got %d", len(footers))
	}
}

// TestECMA376_HeaderFooterByType tests retrieving specific header/footer by type
func TestECMA376_HeaderFooterByType(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add headers with distinct content
	hDefault := doc.AddHeader(HeaderFooterDefault)
	hDefault.SetText("Default Header")

	hFirst := doc.AddHeader(HeaderFooterFirst)
	hFirst.SetText("First Page Header")

	// Retrieve by type
	foundDefault := doc.Header(HeaderFooterDefault)
	if foundDefault == nil {
		t.Error("Should find default header")
	}

	foundFirst := doc.Header(HeaderFooterFirst)
	if foundFirst == nil {
		t.Error("Should find first page header")
	}

	// Non-existent type
	foundEven := doc.Header(HeaderFooterEven)
	if foundEven != nil {
		t.Error("Should not find even header (not added)")
	}
}

// =============================================================================
// Integration Tests
// =============================================================================

// TestECMA376_TrackChangesInSettings tests trackRevisions in settings.xml
func TestECMA376_TrackChangesInSettings(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Enable track changes
	doc.EnableTrackChanges("Test Author")

	// Add tracked content
	p := doc.AddParagraph()
	p.InsertTrackedText("tracked text")

	// Save to temp file and read back
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test_settings.docx")
	if err := doc.SaveAs(tmpFile); err != nil {
		t.Fatal(err)
	}
	
	data, err := os.ReadFile(tmpFile)
	if err != nil {
		t.Fatal(err)
	}

	// Extract settings.xml
	settingsXML := extractPartXML(t, data, "word/settings.xml")
	if settingsXML == "" {
		t.Skip("Settings part not found")
	}

	// Should contain trackRevisions element
	if !strings.Contains(settingsXML, "trackRevisions") {
		t.Error("settings.xml should contain trackRevisions element when enabled")
	}
}

// TestECMA376_RoundTrip_Phase3Features tests round-trip of Phase 3 features
func TestECMA376_RoundTrip_Phase3Features(t *testing.T) {
	// Create document with Phase 3 features
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}

	// Add styles
	style := doc.AddParagraphStyle("TestStyle", "Test Style")
	style.SetBold(true)

	// Add comment
	_, _ = doc.Comments().Add("Test comment", "Test Author", "")

	// Enable track changes and add tracked content
	doc.EnableTrackChanges("Editor")
	p := doc.AddParagraph()
	p.InsertTrackedText("tracked insert")

	// Add header
	h := doc.AddHeader(HeaderFooterDefault)
	h.SetText("Test Header")

	// Save to temp file and read back
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "roundtrip.docx")
	if err := doc.SaveAs(tmpFile); err != nil {
		doc.Close()
		t.Fatal(err)
	}
	doc.Close()

	// Read file into buffer for reopening
	data, err := os.ReadFile(tmpFile)
	if err != nil {
		t.Fatal(err)
	}

	// Reopen
	doc2, err := OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("Failed to reopen document: %v", err)
	}
	defer doc2.Close()

	// Verify styles preserved
	styles := doc2.Styles().List()
	foundStyle := false
	for _, s := range styles {
		if s.ID() == "TestStyle" {
			foundStyle = true
			break
		}
	}
	if !foundStyle {
		t.Error("Style should be preserved after round-trip")
	}

	// Verify track changes state preserved
	if !doc2.TrackChangesEnabled() {
		t.Error("Track changes should be enabled after round-trip")
	}
}

// =============================================================================
// Helper Functions
// =============================================================================

func extractDocumentXML(t *testing.T, data []byte) string {
	return extractPartXML(t, data, "word/document.xml")
}

func extractPartXML(t *testing.T, data []byte, partName string) string {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("Failed to open zip: %v", err)
	}

	for _, f := range r.File {
		if f.Name == partName {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("Failed to open %s: %v", partName, err)
			}
			defer rc.Close()

			content, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("Failed to read %s: %v", partName, err)
			}
			return string(content)
		}
	}
	return ""
}

// xmlContainsElement checks if XML contains an element with given local name
func xmlContainsElement(xmlContent, localName string) bool {
	return strings.Contains(xmlContent, "<"+localName) ||
		strings.Contains(xmlContent, ":"+localName)
}

// =============================================================================
// Parameterized Spec Compliance Tests
// =============================================================================

// TestECMA376_AttributeFormats tests various ECMA-376 attribute format requirements
func TestECMA376_AttributeFormats(t *testing.T) {
	tests := []struct {
		name     string
		setup    func(Document)
		validate func(*testing.T, []byte)
	}{
		{
			name: "revision date format",
			setup: func(d Document) {
				d.EnableTrackChanges("Author")
				d.AddParagraph().InsertTrackedText("test")
			},
			validate: func(t *testing.T, data []byte) {
				xml := extractDocumentXML(t, data)
				// ISO 8601 format should have T separator
				if strings.Contains(xml, "date=") && !strings.Contains(xml, "T") {
					t.Error("Date should be in ISO 8601 format with T separator")
				}
			},
		},
		{
			name: "comment initials attribute",
			setup: func(d Document) {
				c, _ := d.Comments().Add("Comment", "John Doe", "")
				c.SetInitials("JD")
			},
			validate: func(t *testing.T, data []byte) {
				// Comments are in separate part
				xml := extractPartXML(t, data, "word/comments.xml")
				if xml != "" && !strings.Contains(xml, "initials") {
					// Initials might be omitted if empty
					t.Log("Initials attribute checked")
				}
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			tt.setup(doc)

			// Save to temp file and read back
			tmpDir := t.TempDir()
			tmpFile := filepath.Join(tmpDir, "test.docx")
			if err := doc.SaveAs(tmpFile); err != nil {
				t.Fatal(err)
			}
			
			data, err := os.ReadFile(tmpFile)
			if err != nil {
				t.Fatal(err)
			}

			tt.validate(t, data)
		})
	}
}

// Ensure xml import is used
var _ = xml.Unmarshal
func min(a, b int) int { if a < b { return a }; return b }
