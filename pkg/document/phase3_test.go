// Package document provides tests for Phase 3 advanced features.
package document

import (
	"strings"
	"testing"
)

// =============================================================================
// Track Changes Tests
// =============================================================================

func TestEnableDisableTrackChanges(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Initially disabled
	if doc.TrackChangesEnabled() {
		t.Error("Track changes should be disabled by default")
	}

	// Enable with author
	doc.EnableTrackChanges("Test Author")
	if !doc.TrackChangesEnabled() {
		t.Error("Track changes should be enabled")
	}
	if doc.TrackAuthor() != "Test Author" {
		t.Errorf("Expected author 'Test Author', got %q", doc.TrackAuthor())
	}

	// Disable
	doc.DisableTrackChanges()
	if doc.TrackChangesEnabled() {
		t.Error("Track changes should be disabled")
	}
}

func TestSetTrackAuthor(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.SetTrackAuthor("Jane Doe")
	if doc.TrackAuthor() != "Jane Doe" {
		t.Errorf("Expected author 'Jane Doe', got %q", doc.TrackAuthor())
	}
}

func TestInsertTrackedText(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.EnableTrackChanges("Editor")
	p := doc.AddParagraph()
	p.InsertTrackedText("Inserted text")

	// Should have insertion revisions
	insertions := doc.Insertions()
	if len(insertions) != 1 {
		t.Fatalf("Expected 1 insertion, got %d", len(insertions))
	}

	rev := insertions[0]
	if rev.Type() != RevisionInsert {
		t.Error("Expected RevisionInsert type")
	}
	if rev.Author() != "Editor" {
		t.Errorf("Expected author 'Editor', got %q", rev.Author())
	}
	if !strings.Contains(rev.Text(), "Inserted text") {
		t.Errorf("Expected text to contain 'Inserted text', got %q", rev.Text())
	}
}

func TestDeleteTrackedText(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add text without tracking
	p := doc.AddParagraph()
	p.SetText("Original text")

	// Enable tracking and delete
	doc.EnableTrackChanges("Editor")
	err = p.DeleteTrackedText(0)
	if err != nil {
		t.Fatal(err)
	}

	// Should have deletion revision
	deletions := doc.Deletions()
	if len(deletions) != 1 {
		t.Fatalf("Expected 1 deletion, got %d", len(deletions))
	}

	rev := deletions[0]
	if rev.Type() != RevisionDelete {
		t.Error("Expected RevisionDelete type")
	}
}

func TestDeleteTextWithoutTracking(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Add text
	p := doc.AddParagraph()
	p.SetText("Original text")

	// Delete without tracking - should just remove
	err = p.DeleteTrackedText(0)
	if err != nil {
		t.Fatal(err)
	}

	// Should have no deletions
	deletions := doc.Deletions()
	if len(deletions) != 0 {
		t.Errorf("Expected 0 deletions when not tracking, got %d", len(deletions))
	}
}

func TestAcceptAllRevisions(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.EnableTrackChanges("Editor")
	p := doc.AddParagraph()
	p.InsertTrackedText("New text")

	// Should have 1 revision
	if len(doc.AllRevisions()) != 1 {
		t.Fatal("Expected 1 revision")
	}

	// Accept all
	doc.AcceptAllRevisions()

	// Should have 0 revisions after accept
	if len(doc.AllRevisions()) != 0 {
		t.Errorf("Expected 0 revisions after accept, got %d", len(doc.AllRevisions()))
	}
}

func TestRejectAllRevisions(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.EnableTrackChanges("Editor")
	p := doc.AddParagraph()
	p.InsertTrackedText("New text")

	// Should have 1 revision
	if len(doc.AllRevisions()) != 1 {
		t.Fatal("Expected 1 revision")
	}

	// Reject all
	doc.RejectAllRevisions()

	// Should have 0 revisions after reject
	if len(doc.AllRevisions()) != 0 {
		t.Errorf("Expected 0 revisions after reject, got %d", len(doc.AllRevisions()))
	}
}

func TestRevisionDate(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.EnableTrackChanges("Editor")
	p := doc.AddParagraph()
	p.InsertTrackedText("Dated text")

	revisions := doc.AllRevisions()
	if len(revisions) == 0 {
		t.Fatal("Expected at least 1 revision")
	}

	// Date should be set (not zero)
	date := revisions[0].Date()
	if date.IsZero() {
		t.Error("Revision date should not be zero")
	}
}

// =============================================================================
// Comments Tests
// =============================================================================

func TestAddComment(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	c, _ := doc.Comments().Add("This is a comment", "Test Author", "")
	if c == nil {
		t.Fatal("Expected comment to be created")
	}

	if c.Author() != "Test Author" {
		t.Errorf("Expected author 'Test Author', got %q", c.Author())
	}
	if c.Text() != "This is a comment" {
		t.Errorf("Expected text 'This is a comment', got %q", c.Text())
	}
}

func TestCommentInitials(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	c, _ := doc.Comments().Add("Comment text", "John Doe", "")
	c.SetInitials("JD")

	if c.Initials() != "JD" {
		t.Errorf("Expected initials 'JD', got %q", c.Initials())
	}
}

func TestCommentSetText(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	c, _ := doc.Comments().Add("Original", "Author", "")
	c.SetText("Updated comment")

	if c.Text() != "Updated comment" {
		t.Errorf("Expected text 'Updated comment', got %q", c.Text())
	}
}

// =============================================================================
// Content Controls Tests
// =============================================================================

func TestContentControlParagraph(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	cc := p.AddContentControl("Customer", "Customer Name", "Ada Lovelace")
	if cc.Tag() != "Customer" {
		t.Errorf("Tag() = %q, want %q", cc.Tag(), "Customer")
	}
	if cc.Alias() != "Customer Name" {
		t.Errorf("Alias() = %q, want %q", cc.Alias(), "Customer Name")
	}
	if cc.Text() != "Ada Lovelace" {
		t.Errorf("Text() = %q, want %q", cc.Text(), "Ada Lovelace")
	}
}

func TestContentControlBlock(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	cc := doc.AddBlockContentControl("OrderId", "Order ID", "SO-123")
	if cc.Tag() != "OrderId" {
		t.Errorf("Tag() = %q, want %q", cc.Tag(), "OrderId")
	}
	if len(doc.Body().ContentControls()) != 1 {
		t.Error("Expected one block content control")
	}
}

func TestContentControlSetters(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	cc := p.AddContentControl("Customer", "Customer Name", "Ada")
	cc.SetTag("Order")
	cc.SetAlias("Order Name")
	cc.SetContentControlID(42)
	if err := cc.SetContentControlLock("content"); err != nil {
		t.Fatalf("SetContentControlLock() error = %v", err)
	}
	if cc.Tag() != "Order" {
		t.Errorf("Tag() = %q, want %q", cc.Tag(), "Order")
	}
	if cc.Alias() != "Order Name" {
		t.Errorf("Alias() = %q, want %q", cc.Alias(), "Order Name")
	}
	if cc.ID() != 42 {
		t.Errorf("ID() = %d, want %d", cc.ID(), 42)
	}
	if cc.Lock() != "content" {
		t.Errorf("Lock() = %q, want %q", cc.Lock(), "content")
	}
}

func TestContentControlSetTextInline(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	cc := p.AddContentControl("", "", "Old")
	cc.SetText("New")
	if cc.Text() != "New" {
		t.Errorf("Text() = %q, want %q", cc.Text(), "New")
	}
	if !cc.IsInline() {
		t.Error("Expected inline content control")
	}
	if cc.IsBlock() {
		t.Error("Expected inline content control to not be block")
	}
}

func TestContentControlSetTextBlock(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	cc := doc.AddBlockContentControl("Block", "Block Alias", "Old")
	cc.SetText("New")
	if cc.Text() != "New" {
		t.Errorf("Text() = %q, want %q", cc.Text(), "New")
	}
	if !cc.IsBlock() {
		t.Error("Expected block content control")
	}
	if len(cc.Paragraphs()) != 1 {
		t.Errorf("Expected 1 paragraph, got %d", len(cc.Paragraphs()))
	}
}

func TestContentControlCollections(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	p.AddContentControl("InlineTag", "", "Inline")
	doc.AddBlockContentControl("BlockTag", "", "Block")

	all := doc.ContentControls()
	if len(all) != 2 {
		t.Fatalf("Expected 2 content controls, got %d", len(all))
	}
	inline := doc.ContentControlsByTag("InlineTag")
	if len(inline) != 1 {
		t.Fatalf("Expected 1 inline content control, got %d", len(inline))
	}
	block := doc.ContentControlByTag("BlockTag")
	if block == nil {
		t.Fatal("Expected content control by tag")
	}
}

func TestContentControlRemove(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	cc := p.AddContentControl("Remove", "", "Text")
	if err := cc.Remove(); err != nil {
		t.Fatalf("Remove() error = %v", err)
	}
	if len(p.ContentControls()) != 0 {
		t.Error("Expected inline content control to be removed")
	}

	cc2 := doc.AddBlockContentControl("RemoveBlock", "", "Text")
	if err := cc2.Remove(); err != nil {
		t.Fatalf("Remove() error = %v", err)
	}
	if len(doc.Body().ContentControls()) != 0 {
		t.Error("Expected block content control to be removed")
	}
}

func TestContentControlDropDownList(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	cc := doc.AddBlockContentControl("Choices", "Choices", "Select")
	cc.SetDropDownList([]ContentControlListItem{
		{DisplayText: "Option A", Value: "A"},
		{DisplayText: "Option B", Value: "B"},
	})
	items := cc.ListItems()
	if len(items) != 2 {
		t.Fatalf("Expected 2 list items, got %d", len(items))
	}
	if items[0].DisplayText != "Option A" || items[0].Value != "A" {
		t.Errorf("Unexpected first item: %+v", items[0])
	}
	cc.ClearListControl()
	if len(cc.ListItems()) != 0 {
		t.Error("Expected list items to be cleared")
	}
}

func TestContentControlComboBox(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	cc := doc.AddBlockContentControl("Combo", "Combo", "Select")
	cc.SetComboBox([]ContentControlListItem{
		{DisplayText: "One", Value: "1"},
	})
	items := cc.ListItems()
	if len(items) != 1 || items[0].Value != "1" {
		t.Errorf("Unexpected combo box items: %+v", items)
	}
	cc.ClearListControl()
	if len(cc.ListItems()) != 0 {
		t.Error("Expected combo box items to be cleared")
	}
}

func TestContentControlDateConfig(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	cc := doc.AddBlockContentControl("Date", "Date", "2026-02-01")
	cc.SetDateConfig(ContentControlDateConfig{
		Format:   "yyyy-MM-dd",
		Locale:   "en-US",
		Calendar: "gregorian",
		StoreMappedDataAs: "date",
	})
	cfg := cc.DateConfig()
	if cfg == nil {
		t.Fatal("Expected date config")
	}
	if cfg.Format != "yyyy-MM-dd" || cfg.Locale != "en-US" || cfg.Calendar != "gregorian" {
		t.Errorf("Unexpected date config: %+v", cfg)
	}
	cc.ClearDateConfig()
	if cc.DateConfig() != nil {
		t.Error("Expected date config to be cleared")
	}
}

func TestContentControlDropDownListRoundTrip(t *testing.T) {
	h := NewTestHelper(t)
doc := h.CreateDocument(func(d Document) {
		cc := d.AddBlockContentControl("Choices", "Choices", "Select")
		cc.SetDropDownList([]ContentControlListItem{
			{DisplayText: "Option A", Value: "A"},
			{DisplayText: "Option B", Value: "B"},
		})
	})
	defer doc.Close()

	path := h.SaveDocument(doc, "sdt_dropdown_roundtrip.docx")
	doc.Close()

	doc2 := h.OpenDocument(path)
	defer doc2.Close()

	cc := doc2.ContentControlByTag("Choices")
	if cc == nil {
		t.Fatal("Expected content control after round-trip")
	}
	items := cc.ListItems()
	if len(items) != 2 {
		t.Fatalf("Expected 2 list items, got %d", len(items))
	}
	if items[1].DisplayText != "Option B" || items[1].Value != "B" {
		t.Errorf("Unexpected list items after round-trip: %+v", items)
	}
}

func TestContentControlDateConfigRoundTrip(t *testing.T) {
	h := NewTestHelper(t)
doc := h.CreateDocument(func(d Document) {
		cc := d.AddBlockContentControl("Date", "Date", "2026-02-01")
		cc.SetDateConfig(ContentControlDateConfig{
			Format:   "yyyy-MM-dd",
			Locale:   "en-US",
			Calendar: "gregorian",
		})
	})
	defer doc.Close()

	path := h.SaveDocument(doc, "sdt_date_roundtrip.docx")
	doc.Close()

	doc2 := h.OpenDocument(path)
	defer doc2.Close()

	cc := doc2.ContentControlByTag("Date")
	if cc == nil {
		t.Fatal("Expected content control after round-trip")
	}
	cfg := cc.DateConfig()
	if cfg == nil {
		t.Fatal("Expected date config after round-trip")
	}
	if cfg.Format != "yyyy-MM-dd" || cfg.Locale != "en-US" || cfg.Calendar != "gregorian" {
		t.Errorf("Unexpected date config after round-trip: %+v", cfg)
	}
}

// =============================================================================
// Hyperlink/Bookmark Tests
// =============================================================================

func TestAddHyperlink(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	link, err := p.AddHyperlink("https://example.com", "Example")
	if err != nil {
		t.Fatalf("AddHyperlink() error = %v", err)
	}
	if link.URL() != "https://example.com" {
		t.Errorf("URL() = %q, want %q", link.URL(), "https://example.com")
	}
	if link.Text() != "Example" {
		t.Errorf("Text() = %q, want %q", link.Text(), "Example")
	}
}

func TestAddBookmarkAndLink(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	p.AddRun().SetText("Go")
	p.AddRun().SetText("OOXML")
	if err := p.AddBookmark("DocStart", 0, 1); err != nil {
		t.Fatalf("AddBookmark() error = %v", err)
	}
	link, err := p.AddBookmarkLink("DocStart", "Jump")
	if err != nil {
		t.Fatalf("AddBookmarkLink() error = %v", err)
	}
	if link.Anchor() != "DocStart" {
		t.Errorf("Anchor() = %q, want %q", link.Anchor(), "DocStart")
	}
}

// =============================================================================
// Fields Tests
// =============================================================================

func TestAddField(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	field, err := p.AddField("PAGE", "1")
	if err != nil {
		t.Fatalf("AddField() error = %v", err)
	}
	if field.Instruction != "PAGE" {
		t.Errorf("Instruction = %q, want %q", field.Instruction, "PAGE")
	}
	if !strings.Contains(p.Text(), "1") {
		t.Error("Expected field display text in paragraph")
	}
}

func TestHyperlinkRoundTrip(t *testing.T) {
	h := NewTestHelper(t)
doc := h.CreateDocument(func(d Document) {
		p := d.AddParagraph()
		_, err := p.AddHyperlink("https://example.com", "Example")
		if err != nil {
			t.Fatalf("AddHyperlink() error = %v", err)
		}
	})
	defer doc.Close()

	path := h.SaveDocument(doc, "hyperlink_roundtrip.docx")
	doc.Close()

	doc2 := h.OpenDocument(path)
	defer doc2.Close()

	paras := doc2.Paragraphs()
	if len(paras) == 0 {
		t.Fatal("Expected paragraph after round-trip")
	}
	links := paras[0].Hyperlinks()
	if len(links) == 0 {
		t.Fatal("Expected hyperlink after round-trip")
	}
	if links[0].URL() != "https://example.com" {
		t.Errorf("URL() = %q, want %q", links[0].URL(), "https://example.com")
	}
}

func TestContentControlRoundTrip(t *testing.T) {
	h := NewTestHelper(t)
doc := h.CreateDocument(func(d Document) {
		d.AddBlockContentControl("Customer", "Customer Name", "Ada")
	})
	defer doc.Close()

	path := h.SaveDocument(doc, "sdt_roundtrip.docx")
	doc.Close()

	doc2 := h.OpenDocument(path)
	defer doc2.Close()

	ccs := doc2.Body().ContentControls()
	if len(ccs) == 0 {
		t.Fatal("Expected content control after round-trip")
	}
	if ccs[0].Tag() != "Customer" {
		t.Errorf("Tag() = %q, want %q", ccs[0].Tag(), "Customer")
	}
}

func TestTechnicalReportWorkflow(t *testing.T) {
	h := NewTestHelper(t)
doc := h.CreateDocument(func(d Document) {
		d.EnableTrackChanges("Test Author")
		h1 := d.AddParagraph()
		h1.SetStyle("Heading1")
		h1.AddRun().SetText("Technical Report")

		table := d.AddTable(3, 2)
		table.Cell(0, 0).SetText("Customer")
		table.Cell(0, 1).SetText("[CUSTOMER_NAME]")
		table.Cell(1, 0).SetText("Project")
		table.Cell(1, 1).SetText("[PROJECT_NAME]")

		target := table.Cell(0, 1).Paragraphs()[0]
		target.InsertTrackedText("Acme Corp")

	_, _ = d.Comments().Add("Verify customer name with legal", "Reviewer", "Acme Corp")
	})
	defer doc.Close()

	path := h.SaveDocument(doc, "technical_report.docx")
	doc.Close()

	doc2 := h.OpenDocument(path)
	defer doc2.Close()

	if !doc2.TrackChangesEnabled() {
		t.Error("Expected track changes enabled after round-trip")
	}
	if len(doc2.AllRevisions()) == 0 {
		t.Error("Expected revisions after round-trip")
	}
	if len(doc2.Comments().All()) != 1 {
		t.Errorf("Expected 1 comment after round-trip, got %d", len(doc2.Comments().All()))
	}
	tables := doc2.Tables()
	if len(tables) != 1 {
		t.Fatalf("Expected 1 table after round-trip, got %d", len(tables))
	}
	if !strings.Contains(tables[0].Cell(0, 1).Text(), "Acme Corp") {
		t.Error("Expected customer cell text after round-trip")
	}
}

func TestContentControlsFixture(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.OpenFixture("sdt_content_controls.docx")
	defer doc.Close()

	ccs := doc.ContentControls()
	if len(ccs) < 3 {
		t.Fatalf("Expected content controls in fixture, got %d", len(ccs))
	}
	inline := doc.ContentControlByTag("InlineTag")
	if inline == nil {
		t.Fatal("Expected inline content control in fixture")
	}
	block := doc.ContentControlByTag("BlockTag")
	if block == nil {
		t.Fatal("Expected block content control in fixture")
	}
	if len(block.ListItems()) != 2 {
		t.Errorf("Expected 2 list items in block control, got %d", len(block.ListItems()))
	}
	date := doc.ContentControlByTag("DateTag")
	if date == nil {
		t.Fatal("Expected date content control in fixture")
	}
	cfg := date.DateConfig()
	if cfg == nil {
		t.Error("Expected date config in fixture")
	}
}

func TestComments(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Initially no comments
	if len(doc.Comments().All()) != 0 {
		t.Error("Expected 0 comments initially")
	}

	// Add multiple comments
	_, _ = doc.Comments().Add("Comment 1", "Author 1", "")
	_, _ = doc.Comments().Add("Comment 2", "Author 2", "")
	_, _ = doc.Comments().Add("Comment 3", "Author 3", "")

	comments := doc.Comments().All()
	if len(comments) != 3 {
		t.Errorf("Expected 3 comments, got %d", len(comments))
	}
}

func TestCommentRepliesRoundTrip(t *testing.T) {
	h := NewTestHelper(t)
	round := h.RoundTrip("comment-replies.docx", func(d Document) {
		c, err := d.Comments().Add("Parent comment", "Author", "")
		if err != nil {
			t.Fatalf("Add() error = %v", err)
		}
		if _, err := c.AddReply("Reply comment", "Reviewer"); err != nil {
			t.Fatalf("AddReply() error = %v", err)
		}
	})

	defer round.Close()

	if len(round.Comments().All()) < 2 {
		t.Fatalf("Expected at least 2 comments, got %d", len(round.Comments().All()))
	}
	foundReply := false
	for _, comment := range round.Comments().All() {
		for _, reply := range comment.Replies() {
			if reply.Text() == "Reply comment" {
				foundReply = true
				break
			}
		}
	}
	if !foundReply {
		t.Error("Expected reply after round-trip")
	}
}

func TestCommentByID(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	c1, _ := doc.Comments().Add("Comment 1", "Author", "")
	c2, _ := doc.Comments().Add("Comment 2", "Author", "")

	// Find by ID
	found := doc.CommentByID(c1.ID())
	if found == nil {
		t.Fatal("Expected to find comment")
	}
	if found.Text() != "Comment 1" {
		t.Errorf("Found wrong comment: %q", found.Text())
	}

	found2 := doc.CommentByID(c2.ID())
	if found2 == nil {
		t.Fatal("Expected to find comment 2")
	}
	if found2.Text() != "Comment 2" {
		t.Errorf("Found wrong comment: %q", found2.Text())
	}

	// Non-existent ID
	notFound := doc.CommentByID("9999")
	if notFound != nil {
		t.Error("Expected nil for non-existent ID")
	}
}

func TestDeleteComment(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	c, _ := doc.Comments().Add("To delete", "Author", "")
	id := c.ID()

	// Delete it
	err = doc.DeleteComment(id)
	if err != nil {
		t.Errorf("DeleteComment failed: %v", err)
	}

	// Should no longer exist
	if doc.CommentByID(id) != nil {
		t.Error("Comment should have been deleted")
	}

	// Delete non-existent - should error
	err = doc.DeleteComment("9999")
	if err == nil {
		t.Error("Expected DeleteComment to error for non-existent")
	}
}

// =============================================================================
// Styles Tests
// =============================================================================

func TestAddParagraphStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	s := doc.AddParagraphStyle("CustomPara", "Custom Paragraph")
	if s == nil {
		t.Fatal("Expected style to be created")
	}

	if s.ID() != "CustomPara" {
		t.Errorf("Expected ID 'CustomPara', got %q", s.ID())
	}
	if s.Name() != "Custom Paragraph" {
		t.Errorf("Expected name 'Custom Paragraph', got %q", s.Name())
	}
	if s.Type() != StyleTypeParagraph {
		t.Errorf("Expected type StyleTypeParagraph, got %v", s.Type())
	}
}

func TestAddCharacterStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	s := doc.AddCharacterStyle("CustomChar", "Custom Character")
	if s.Type() != StyleTypeCharacter {
		t.Errorf("Expected type StyleTypeCharacter, got %v", s.Type())
	}
}

func TestAddTableStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	s := doc.AddTableStyle("CustomTable", "Custom Table")
	if s.Type() != StyleTypeTable {
		t.Errorf("Expected type StyleTypeTable, got %v", s.Type())
	}
}

func TestStyleByID(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.AddParagraphStyle("Style1", "Style One")
	doc.AddParagraphStyle("Style2", "Style Two")

	s := doc.StyleByID("Style1")
	if s == nil {
		t.Fatal("Expected to find style")
	}
	if s.Name() != "Style One" {
		t.Errorf("Found wrong style: %q", s.Name())
	}

	// Non-existent
	notFound := doc.StyleByID("NotExists")
	if notFound != nil {
		t.Error("Expected nil for non-existent style")
	}
}

func TestStyleByName(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.AddParagraphStyle("Style1", "Named Style")

	s := doc.StyleByName("Named Style")
	if s == nil {
		t.Fatal("Expected to find style")
	}
	if s.ID() != "Style1" {
		t.Errorf("Found wrong style: %q", s.ID())
	}
}

func TestDeleteStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.AddParagraphStyle("ToDelete", "Delete Me")

	deleted := doc.DeleteStyle("ToDelete")
	if !deleted {
		t.Error("Expected DeleteStyle to return true")
	}

	if doc.StyleByID("ToDelete") != nil {
		t.Error("Style should have been deleted")
	}

	// Delete non-existent
	deleted = doc.DeleteStyle("NotExists")
	if deleted {
		t.Error("Expected DeleteStyle to return false for non-existent")
	}
}

func TestStyleFormatting(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	s := doc.AddParagraphStyle("Formatted", "Formatted Style")

	// Set formatting
	s.SetBold(true)
	s.SetItalic(true)
	s.SetFontSize(14.0)
	s.SetFontName("Arial")
	s.SetColor("FF0000")
	s.SetAlignment("center")
	s.SetSpacingBefore(240)
	s.SetSpacingAfter(120)
	s.SetNext("NextStyle")
	s.SetLink("LinkedStyle")
	s.SetUIPriority(5)
	s.SetQFormat(true)
	s.SetCustomStyle(true)

	// Verify properties were set (basic check - they don't error)
	if s.ParagraphProperties() == nil {
		t.Error("Expected paragraph properties to be set")
	}
	if s.RunProperties() == nil {
		t.Error("Expected run properties to be set")
	}
	if s.Next() != "NextStyle" {
		t.Errorf("Next() = %q, want NextStyle", s.Next())
	}
	if s.Link() != "LinkedStyle" {
		t.Errorf("Link() = %q, want LinkedStyle", s.Link())
	}
	if s.UIPriority() != 5 {
		t.Errorf("UIPriority() = %d, want 5", s.UIPriority())
	}
	if !s.QFormat() {
		t.Error("QFormat() should be true")
	}
	if !s.CustomStyle() {
		t.Error("CustomStyle() should be true")
	}
}

func TestStyleBasedOn(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.AddParagraphStyle("BaseStyle", "Base")
	s := doc.AddParagraphStyle("DerivedStyle", "Derived")
	s.SetBasedOn("BaseStyle")

	if s.BasedOn() != "BaseStyle" {
		t.Errorf("Expected BasedOn 'BaseStyle', got %q", s.BasedOn())
	}

	// Clear BasedOn
	s.SetBasedOn("")
	if s.BasedOn() != "" {
		t.Error("Expected empty BasedOn after clearing")
	}
}

func TestStyleDefault(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	s := doc.AddParagraphStyle("DefaultPara", "Default Paragraph")
	s.SetDefault(true)

	if !s.IsDefault() {
		t.Error("Expected style to be default")
	}

	// Check DefaultParagraphStyle returns it
	defaultStyle := doc.DefaultParagraphStyle()
	if defaultStyle == nil {
		t.Fatal("Expected to find default paragraph style")
	}
	if defaultStyle.ID() != "DefaultPara" {
		t.Errorf("Wrong default style: %q", defaultStyle.ID())
	}
}

func TestStyles(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	// Initially no styles
	if len(doc.Styles().List()) != 0 {
		t.Error("Expected 0 styles initially")
	}

	// Add styles
	doc.AddParagraphStyle("Para1", "Paragraph 1")
	doc.AddCharacterStyle("Char1", "Character 1")
	doc.AddTableStyle("Table1", "Table 1")
	doc.AddNumberingStyle("Num1", "Numbering 1")

	styles := doc.Styles().List()
	if len(styles) != 4 {
		t.Errorf("Expected 4 styles, got %d", len(styles))
	}
}

// =============================================================================
// Numbering Tests
// =============================================================================

func TestAddNumberedListStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	numID, err := doc.AddNumberedListStyle()
	if err != nil {
		t.Fatalf("AddNumberedListStyle() error = %v", err)
	}
	if numID == 0 {
		t.Error("Expected non-zero numbering ID")
	}

	p := doc.AddParagraph()
	if err := p.SetList(numID, 0); err != nil {
		t.Fatalf("SetList() error = %v", err)
	}

	if p.ListNumberingID() != numID {
		t.Errorf("ListNumberingID() = %d, want %d", p.ListNumberingID(), numID)
	}
	if p.ListLevel() != 0 {
		t.Errorf("ListLevel() = %d, want 0", p.ListLevel())
	}
}

func TestNumberingRoundTrip(t *testing.T) {
	h := NewTestHelper(t)
doc := h.CreateDocument(func(d Document) {
		numID, err := d.AddNumberedListStyle()
		if err != nil {
			t.Fatalf("AddNumberedListStyle() error = %v", err)
		}
		p := d.AddParagraph()
		p.SetText("Item 1")
		if err := p.SetList(numID, 0); err != nil {
			t.Fatalf("SetList() error = %v", err)
		}
	})
	defer doc.Close()

	path := h.SaveDocument(doc, "numbering_roundtrip.docx")
	doc.Close()

	doc2 := h.OpenDocument(path)
	defer doc2.Close()

	if len(doc2.Numbering()) == 0 {
		t.Fatal("Expected numbering definitions after round-trip")
	}
	paras := doc2.Paragraphs()
	if len(paras) == 0 {
		t.Fatal("Expected paragraph after round-trip")
	}
	if paras[0].ListNumberingID() == 0 {
		t.Error("Expected numbering ID on paragraph after round-trip")
	}
}

// =============================================================================
// Headers/Footers Tests
// =============================================================================

func TestAddHeader(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	h := doc.AddHeader(HeaderFooterDefault)
	if h == nil {
		t.Fatal("Expected header to be created")
	}
	if h.Type() != HeaderFooterDefault {
		t.Errorf("Expected type 'default', got %v", h.Type())
	}
}

func TestAddFooter(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	f := doc.AddFooter(HeaderFooterDefault)
	if f == nil {
		t.Fatal("Expected footer to be created")
	}
	if f.Type() != HeaderFooterDefault {
		t.Errorf("Expected type 'default', got %v", f.Type())
	}
}

func TestHeaderAddParagraph(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	h := doc.AddHeader(HeaderFooterDefault)
	p := h.AddParagraph()
	if p == nil {
		t.Fatal("Expected paragraph to be added")
	}
}

func TestHeaderSetText(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	h := doc.AddHeader(HeaderFooterDefault)
	h.SetText("Header Text")

	if !strings.Contains(h.Text(), "Header Text") {
		t.Errorf("Expected header to contain 'Header Text', got %q", h.Text())
	}
}

func TestFooterSetText(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	f := doc.AddFooter(HeaderFooterDefault)
	f.SetText("Footer Text")

	if !strings.Contains(f.Text(), "Footer Text") {
		t.Errorf("Expected footer to contain 'Footer Text', got %q", f.Text())
	}
}

func TestMultipleHeaderTypes(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.AddHeader(HeaderFooterDefault)
	doc.AddHeader(HeaderFooterFirst)
	doc.AddHeader(HeaderFooterEven)

	headers := doc.Headers()
	if len(headers) != 3 {
		t.Errorf("Expected 3 headers, got %d", len(headers))
	}
}

func TestHeaderFooterGetByType(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()

	doc.AddHeader(HeaderFooterDefault)
	doc.AddHeader(HeaderFooterFirst)

	h := doc.Header(HeaderFooterDefault)
	if h == nil {
		t.Error("Expected to find default header")
	}

	h2 := doc.Header(HeaderFooterFirst)
	if h2 == nil {
		t.Error("Expected to find first-page header")
	}

	// Non-existent type
	h3 := doc.Header(HeaderFooterEven)
	if h3 != nil {
		t.Error("Expected nil for non-existent header type")
	}
}

// =============================================================================
// Parameterized Tests
// =============================================================================

func TestRevisionTypes(t *testing.T) {
	tests := []struct {
		name     string
		revType  RevisionType
		expected string
	}{
		{"Insert", RevisionInsert, "insert"},
		{"Delete", RevisionDelete, "delete"},
		{"Format", RevisionFormat, "format"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := tt.revType.String()
			if got != tt.expected {
				t.Errorf("RevisionType.String() = %q, want %q", got, tt.expected)
			}
		})
	}
}

func TestStyleTypes(t *testing.T) {
	tests := []struct {
		name      string
		styleType StyleType
	}{
		{"Paragraph", StyleTypeParagraph},
		{"Character", StyleTypeCharacter},
		{"Table", StyleTypeTable},
		{"Numbering", StyleTypeNumbering},
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
				s = doc.AddParagraphStyle("Test"+tt.name, "Test "+tt.name)
			case StyleTypeCharacter:
				s = doc.AddCharacterStyle("Test"+tt.name, "Test "+tt.name)
			case StyleTypeTable:
				s = doc.AddTableStyle("Test"+tt.name, "Test "+tt.name)
			case StyleTypeNumbering:
				s = doc.AddNumberingStyle("Test"+tt.name, "Test "+tt.name)
			}

			if s != nil && s.Type() != tt.styleType {
				t.Errorf("Style type = %v, want %v", s.Type(), tt.styleType)
			}
		})
	}
}

func TestHeaderFooterTypes(t *testing.T) {
	tests := []struct {
		name   string
		hfType HeaderFooterType
	}{
		{"Default", HeaderFooterDefault},
		{"First", HeaderFooterFirst},
		{"Even", HeaderFooterEven},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()

			h := doc.AddHeader(tt.hfType)
			if h.Type() != tt.hfType {
				t.Errorf("Header type = %v, want %v", h.Type(), tt.hfType)
			}

			f := doc.AddFooter(tt.hfType)
			if f.Type() != tt.hfType {
				t.Errorf("Footer type = %v, want %v", f.Type(), tt.hfType)
			}
		})
	}
}
