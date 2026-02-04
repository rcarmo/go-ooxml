// Package document provides consolidated parameterized tests using the test framework.
// This file demonstrates best practices for test organization and reduces duplication.
package document

import (
	"strings"
	"testing"
)

// =============================================================================
// Phase 3 Fixture Round-Trip Tests
// =============================================================================

// TestPhase3Fixtures_RoundTrip tests all Phase 3 fixtures survive round-trip.
func TestPhase3Fixtures_RoundTrip(t *testing.T) {
	RunFixtureTests(t, Phase3Fixtures, func(t *testing.T, h *TestHelper, fixture TestFixture) {
		// Create and save
		doc := h.CreateDocument(fixture.Setup)
		path := h.SaveDocument(doc, fixture.Name+".docx")
		doc.Close()
		
		// Reopen - should not error
		doc2 := h.OpenDocument(path)
		defer doc2.Close()
		
		// Basic structural check
		if doc2 == nil {
			t.Error("Reopened document should not be nil")
		}
	})
}

// =============================================================================
// Parameterized Track Changes Tests
// =============================================================================

// TestTrackChanges_Parameterized tests track changes with various inputs.
func TestTrackChanges_Parameterized(t *testing.T) {
	RunTrackChangesTests(t, TrackChangesTestCases(), func(t *testing.T, h *TestHelper, tc TrackChangesTestCase) {
		doc := h.CreateDocument(func(d Document) {
			d.EnableTrackChanges(tc.Author)
			p := d.AddParagraph()
			p.InsertTrackedText(tc.Text)
		})
		defer doc.Close()
		
		// Check enabled state
		h.AssertTrackChangesEnabled(doc, tc.ExpectEnabled)
		
		// Check author
		h.AssertTrackAuthor(doc, tc.Author)
		
		// Check revision exists
		if tc.ExpectRevision {
			h.AssertInsertionCount(doc, 1)
		}
	})
}

// TestTrackChanges_RoundTrip tests track changes survive save/reopen.
func TestTrackChanges_RoundTrip(t *testing.T) {
	RunTrackChangesTests(t, TrackChangesTestCases(), func(t *testing.T, h *TestHelper, tc TrackChangesTestCase) {
		doc := h.RoundTrip(tc.Name+".docx", func(d Document) {
			d.EnableTrackChanges(tc.Author)
			d.AddParagraph().InsertTrackedText(tc.Text)
		})
		defer doc.Close()
		
		// Track changes state should be preserved
		h.AssertTrackChangesEnabled(doc, tc.ExpectEnabled)
		
		// Revisions should be preserved
		if tc.ExpectRevision {
			revisions := doc.AllRevisions()
			if len(revisions) == 0 {
				t.Error("Expected revisions to survive round-trip")
			}
		}
	})
}

// TestTrackChanges_AcceptReject tests accepting and rejecting revisions.
func TestTrackChanges_AcceptReject(t *testing.T) {
	tests := []struct {
		name     string
		action   string // "accept" or "reject"
		wantLeft int    // expected revision count after action
	}{
		{"accept all", "accept", 0},
		{"reject all", "reject", 0},
	}
	
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d Document) {
				d.EnableTrackChanges("Editor")
				d.AddParagraph().InsertTrackedText("Insert 1")
				d.AddParagraph().InsertTrackedText("Insert 2")
			})
			defer doc.Close()
			
			// Verify initial count
			initialCount := len(doc.AllRevisions())
			if initialCount != 2 {
				t.Fatalf("Expected 2 initial revisions, got %d", initialCount)
			}
			
			// Apply action
			switch tt.action {
			case "accept":
				doc.AcceptAllRevisions()
			case "reject":
				doc.RejectAllRevisions()
			}
			
			// Verify result
			h.AssertRevisionCount(doc, tt.wantLeft)
		})
	}
}

// =============================================================================
// Parameterized Comment Tests
// =============================================================================

// TestComments_Parameterized tests comments with various inputs.
func TestComments_Parameterized(t *testing.T) {
	RunCommentTests(t, CommentTestCases(), func(t *testing.T, h *TestHelper, tc CommentTestCase) {
		doc := h.CreateDocument(func(d Document) {
			d.AddParagraph().SetText("Document text")
			c, _ := d.Comments().Add(tc.Text, tc.Author, "")
			if tc.Initials != "" {
				c.SetInitials(tc.Initials)
			}
		})
		defer doc.Close()
		
		// Check comment count
		h.AssertCommentCount(doc, 1)
		
		// Check comment properties
		c := doc.Comments().All()[0]
		if c.Author() != tc.Author {
			t.Errorf("Author() = %q, want %q", c.Author(), tc.Author)
		}
		if c.Text() != tc.Text {
			t.Errorf("Text() = %q, want %q", c.Text(), tc.Text)
		}
		if tc.Initials != "" && c.Initials() != tc.Initials {
			t.Errorf("Initials() = %q, want %q", c.Initials(), tc.Initials)
		}
	})
}

// TestComments_RoundTrip tests comments survive save/reopen.
func TestComments_RoundTrip(t *testing.T) {
	RunCommentTests(t, CommentTestCases(), func(t *testing.T, h *TestHelper, tc CommentTestCase) {
		doc := h.RoundTrip(tc.Name+".docx", func(d Document) {
			d.AddParagraph().SetText("Text")
			c, _ := d.Comments().Add(tc.Text, tc.Author, "")
			if tc.Initials != "" {
				c.SetInitials(tc.Initials)
			}
		})
		defer doc.Close()
		
		// Comments should survive
		comments := doc.Comments().All()
		if len(comments) == 0 {
			t.Error("Expected comments to survive round-trip")
			return
		}
		
		// Check content preserved
		c := comments[0]
		if c.Author() != tc.Author {
			t.Errorf("Author after round-trip = %q, want %q", c.Author(), tc.Author)
		}
	})
}

// TestComments_CRUD tests create, read, update, delete operations.
func TestComments_CRUD(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()
	
	// Create
	c1, _ := doc.Comments().Add("Comment 1", "Author1", "")
	c2, _ := doc.Comments().Add("Comment 2", "Author2", "")
	h.AssertCommentCount(doc, 2)
	
	// Read
	found := doc.CommentByID(c1.ID())
	if found == nil {
		t.Fatal("Should find comment by ID")
	}
	if found.Text() != "Comment 1" {
		t.Errorf("Found wrong comment")
	}
	
	// Update
	c1.SetText("Updated Comment 1")
	if c1.Text() != "Updated Comment 1" {
		t.Error("Text update failed")
	}
	
	// Delete
	doc.DeleteComment(c2.ID())
	h.AssertCommentCount(doc, 1)
	
	// Verify correct one was deleted
	if doc.CommentByID(c2.ID()) != nil {
		t.Error("Deleted comment should not be found")
	}
	if doc.CommentByID(c1.ID()) == nil {
		t.Error("Remaining comment should still exist")
	}
}

// =============================================================================
// Parameterized Style Tests
// =============================================================================

// TestStyles_Parameterized tests styles with various configurations.
func TestStyles_Parameterized(t *testing.T) {
	RunStyleTests(t, StyleTestCases(), func(t *testing.T, h *TestHelper, tc StyleTestCase) {
		doc := h.CreateDocument(func(d Document) {
			var s Style
			switch tc.StyleType {
			case StyleTypeParagraph:
				s = d.AddParagraphStyle(tc.StyleID, tc.StyleName)
			case StyleTypeCharacter:
				s = d.AddCharacterStyle(tc.StyleID, tc.StyleName)
			case StyleTypeTable:
				s = d.AddTableStyle(tc.StyleID, tc.StyleName)
			case StyleTypeNumbering:
				s = d.AddNumberingStyle(tc.StyleID, tc.StyleName)
			default:
				t.Skipf("Style type %v not implemented", tc.StyleType)
				return
			}
			
			if tc.Bold {
				s.SetBold(true)
			}
			if tc.Italic {
				s.SetItalic(true)
			}
			if tc.FontSize > 0 {
				s.SetFontSize(tc.FontSize)
			}
			if tc.Color != "" {
				s.SetColor(tc.Color)
			}
		})
		defer doc.Close()
		
		// Check style exists
		h.AssertStyleCount(doc, 1)
		
		// Check style properties
		s := doc.StyleByID(tc.StyleID)
		if s == nil {
			t.Fatalf("Style %q not found", tc.StyleID)
		}
		if s.Name() != tc.StyleName {
			t.Errorf("Name() = %q, want %q", s.Name(), tc.StyleName)
		}
		if s.Type() != tc.StyleType {
			t.Errorf("Type() = %v, want %v", s.Type(), tc.StyleType)
		}
	})
}

// TestStyles_Inheritance tests style inheritance (basedOn).
func TestStyles_Inheritance(t *testing.T) {
	tests := []struct {
		name     string
		baseID   string
		childID  string
		wantLink string
	}{
		{"single level", "Base", "Child", "Base"},
		{"different names", "Parent", "Derived", "Parent"},
	}
	
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d Document) {
				d.AddParagraphStyle(tt.baseID, "Base Style")
				child := d.AddParagraphStyle(tt.childID, "Child Style")
				child.SetBasedOn(tt.baseID)
			})
			defer doc.Close()
			
			child := doc.StyleByID(tt.childID)
			if child == nil {
				t.Fatal("Child style not found")
			}
			if child.BasedOn() != tt.wantLink {
				t.Errorf("BasedOn() = %q, want %q", child.BasedOn(), tt.wantLink)
			}
		})
	}
}

// =============================================================================
// Parameterized Header/Footer Tests
// =============================================================================

// TestHeadersFooters_Parameterized tests headers/footers with various types.
func TestHeadersFooters_Parameterized(t *testing.T) {
	RunHeaderFooterTests(t, HeaderFooterTestCases(), func(t *testing.T, h *TestHelper, tc HeaderFooterTestCase) {
		doc := h.CreateDocument(func(d Document) {
			if tc.IsHeader {
				hdr := d.AddHeader(tc.Type)
				hdr.SetText(tc.Text)
			} else {
				ftr := d.AddFooter(tc.Type)
				ftr.SetText(tc.Text)
			}
		})
		defer doc.Close()
		
		// Check count
		if tc.IsHeader {
			h.AssertHeaderCount(doc, 1)
		} else {
			h.AssertFooterCount(doc, 1)
		}
		
		// Check retrieval by type
		if tc.IsHeader {
			hdr := doc.Header(tc.Type)
			if hdr == nil {
				t.Fatalf("Header type %v not found", tc.Type)
			}
			if !strings.Contains(hdr.Text(), tc.Text) {
				t.Errorf("Header text = %q, should contain %q", hdr.Text(), tc.Text)
			}
		} else {
			ftr := doc.Footer(tc.Type)
			if ftr == nil {
				t.Fatalf("Footer type %v not found", tc.Type)
			}
			if !strings.Contains(ftr.Text(), tc.Text) {
				t.Errorf("Footer text = %q, should contain %q", ftr.Text(), tc.Text)
			}
		}
	})
}

// =============================================================================
// Combinatorial Formatting Tests
// =============================================================================

// TestFormatCombinations_Subset tests a subset of format combinations.
// Full combinatorial testing would be 216 combinations - we test a representative subset.
func TestFormatCombinations_Subset(t *testing.T) {
	// Select representative combinations
	combinations := []FormatCombination{
		{Bold: false, Italic: false, Underline: false, FontSize: 11, Color: ""},
		{Bold: true, Italic: false, Underline: false, FontSize: 11, Color: ""},
		{Bold: false, Italic: true, Underline: false, FontSize: 11, Color: ""},
		{Bold: false, Italic: false, Underline: true, FontSize: 11, Color: ""},
		{Bold: true, Italic: true, Underline: false, FontSize: 14, Color: ""},
		{Bold: true, Italic: true, Underline: true, FontSize: 18, Color: "FF0000"},
		{Bold: false, Italic: false, Underline: false, FontSize: 18, Color: "0000FF"},
	}
	
	for _, fc := range combinations {
		t.Run(fc.String(), func(t *testing.T) {
			h := NewTestHelper(t)
			
			// Use a sanitized filename (no slashes)
			filename := strings.ReplaceAll(fc.String(), "/", "_") + ".docx"
			
			// Create document with formatted text
			doc := h.RoundTrip(filename, func(d Document) {
				p := d.AddParagraph()
				r := p.AddRun()
				r.SetText("Formatted text")
				fc.ApplyToRun(r)
			})
			defer doc.Close()
			
			// Basic verification - document should survive round-trip
			paras := doc.Paragraphs()
			if len(paras) == 0 {
				t.Error("Expected at least one paragraph")
				return
			}
			
			// Text should be preserved
			text := paras[0].Text()
			if !strings.Contains(text, "Formatted") {
				t.Errorf("Text not preserved: %q", text)
			}
		})
	}
}

// =============================================================================
// Integration Tests - Combined Features
// =============================================================================

// TestIntegration_FullDocument tests all Phase 3 features together.
func TestIntegration_FullDocument(t *testing.T) {
	h := NewTestHelper(t)
	
	// Create document with all features
	doc := h.CreateDocument(func(d Document) {
		// 1. Add styles
		titleStyle := d.AddParagraphStyle("DocTitle", "Document Title")
		titleStyle.SetBold(true)
		titleStyle.SetFontSize(24)
		
		bodyStyle := d.AddParagraphStyle("DocBody", "Body Text")
		bodyStyle.SetFontSize(11)
		
		// 2. Add header/footer
		d.AddHeader(HeaderFooterDefault).SetText("Header Text")
		d.AddFooter(HeaderFooterDefault).SetText("Footer Text")
		
		// 3. Add content with styles
		title := d.AddParagraph()
		title.SetStyle("DocTitle")
		title.SetText("Document Title")
		
		body := d.AddParagraph()
		body.SetStyle("DocBody")
		body.SetText("Body text content")
		
		// 4. Enable track changes and add more content
		d.EnableTrackChanges("Editor")
		tracked := d.AddParagraph()
		tracked.InsertTrackedText("Tracked insertion")
		
		// 5. Add comments
		_, _ = d.Comments().Add("Please review", "Reviewer", "")
	})
	
	// Verify all features
	h.AssertStyleCount(doc, 2)
	h.AssertHeaderCount(doc, 1)
	h.AssertFooterCount(doc, 1)
	h.AssertTrackChangesEnabled(doc, true)
	h.AssertInsertionCount(doc, 1)
	h.AssertCommentCount(doc, 1)
	h.AssertParagraphCount(doc, 3)
	
	// Save and verify persistence
	path := h.SaveDocument(doc, "integration_full.docx")
	doc.Close()
	
	doc2 := h.OpenDocument(path)
	defer doc2.Close()
	
	// Verify after round-trip
	h.AssertParagraphCount(doc2, 3)
	h.AssertTrackChangesEnabled(doc2, true)
}

// TestIntegration_MultipleEditors tests document with multiple editors.
func TestIntegration_MultipleEditors(t *testing.T) {
	h := NewTestHelper(t)
	
	editors := []string{"Alice", "Bob", "Charlie"}
	
	doc := h.CreateDocument(func(d Document) {
		for i, editor := range editors {
			d.EnableTrackChanges(editor)
			p := d.AddParagraph()
			p.InsertTrackedText("Edit from " + editor)
			_, _ = d.Comments().Add("Comment from " + editor, editor, "")
			
			// Verify current author
			if d.TrackAuthor() != editor {
				t.Errorf("Iteration %d: TrackAuthor() = %q, want %q", i, d.TrackAuthor(), editor)
			}
		}
	})
	defer doc.Close()
	
	// Verify all edits present
	h.AssertInsertionCount(doc, len(editors))
	h.AssertCommentCount(doc, len(editors))
	
	// Verify different authors in revisions
	revisions := doc.Insertions()
	authorsSeen := make(map[string]bool)
	for _, rev := range revisions {
		authorsSeen[rev.Author()] = true
	}
	if len(authorsSeen) != len(editors) {
		t.Errorf("Expected %d different authors, got %d", len(editors), len(authorsSeen))
	}
}
