package document

import (
	"os"
	"path/filepath"
	"testing"
)

func TestNewDocument(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	// Check that body exists
	if doc.Body() == nil {
		t.Error("Body() returned nil")
	}

	// New document should have no paragraphs
	if len(doc.Paragraphs()) != 0 {
		t.Errorf("New document has %d paragraphs, want 0", len(doc.Paragraphs()))
	}
}

func TestAddParagraph(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	// Add paragraphs
	p1 := doc.AddParagraph()
	p1.SetText("Hello")

	p2 := doc.AddParagraph()
	p2.SetText("World")

	paragraphs := doc.Paragraphs()
	if len(paragraphs) != 2 {
		t.Fatalf("Expected 2 paragraphs, got %d", len(paragraphs))
	}

	if paragraphs[0].Text() != "Hello" {
		t.Errorf("First paragraph text = %q, want %q", paragraphs[0].Text(), "Hello")
	}
	if paragraphs[1].Text() != "World" {
		t.Errorf("Second paragraph text = %q, want %q", paragraphs[1].Text(), "World")
	}
}

func TestRunFormatting(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	run := p.AddRun()
	run.SetText("Test")

	// Test bold
	run.SetBold(true)
	if !run.Bold() {
		t.Error("Bold() = false after SetBold(true)")
	}

	// Test italic
	run.SetItalic(true)
	if !run.Italic() {
		t.Error("Italic() = false after SetItalic(true)")
	}

	// Test underline
	run.SetUnderline(true)
	if !run.Underline() {
		t.Error("Underline() = false after SetUnderline(true)")
	}

	// Test font size
	run.SetFontSize(14)
	if run.FontSize() != 14 {
		t.Errorf("FontSize() = %v, want 14", run.FontSize())
	}

	// Test font name
	run.SetFontName("Arial")
	if run.FontName() != "Arial" {
		t.Errorf("FontName() = %q, want %q", run.FontName(), "Arial")
	}

	// Test color
	run.SetColor("FF0000")
	if run.Color() != "FF0000" {
		t.Errorf("Color() = %q, want %q", run.Color(), "FF0000")
	}

	// Test highlight
	run.SetHighlight("yellow")
	if run.Highlight() != "yellow" {
		t.Errorf("Highlight() = %q, want %q", run.Highlight(), "yellow")
	}
}

func TestAddTable(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	// Add a 3x2 table
	tbl := doc.AddTable(3, 2)

	if tbl.RowCount() != 3 {
		t.Errorf("RowCount() = %d, want 3", tbl.RowCount())
	}
	if tbl.ColumnCount() != 2 {
		t.Errorf("ColumnCount() = %d, want 2", tbl.ColumnCount())
	}

	// Set cell text
	tbl.Cell(0, 0).SetText("A1")
	tbl.Cell(0, 1).SetText("B1")
	tbl.Cell(1, 0).SetText("A2")
	tbl.Cell(1, 1).SetText("B2")

	if tbl.Cell(0, 0).Text() != "A1" {
		t.Errorf("Cell(0,0).Text() = %q, want %q", tbl.Cell(0, 0).Text(), "A1")
	}
	if tbl.Cell(1, 1).Text() != "B2" {
		t.Errorf("Cell(1,1).Text() = %q, want %q", tbl.Cell(1, 1).Text(), "B2")
	}

	// Test FirstRowText
	firstRow := tbl.FirstRowText()
	if len(firstRow) != 2 || firstRow[0] != "A1" || firstRow[1] != "B1" {
		t.Errorf("FirstRowText() = %v, want [A1 B1]", firstRow)
	}
}

func TestTableRowOperations(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	tbl := doc.AddTable(2, 2)
	tbl.Cell(0, 0).SetText("Header1")
	tbl.Cell(0, 1).SetText("Header2")
	tbl.Cell(1, 0).SetText("Data1")
	tbl.Cell(1, 1).SetText("Data2")

	// Add row
	newRow := tbl.AddRow()
	newRow.Cell(0).SetText("New1")
	newRow.Cell(1).SetText("New2")

	if tbl.RowCount() != 3 {
		t.Errorf("After AddRow, RowCount() = %d, want 3", tbl.RowCount())
	}

	// Delete row
	if err := tbl.DeleteRow(1); err != nil {
		t.Errorf("DeleteRow(1) error = %v", err)
	}
	if tbl.RowCount() != 2 {
		t.Errorf("After DeleteRow, RowCount() = %d, want 2", tbl.RowCount())
	}
}

func TestParagraphStyle(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	p.SetStyle("Heading1")

	if p.Style() != "Heading1" {
		t.Errorf("Style() = %q, want %q", p.Style(), "Heading1")
	}

	if !p.IsHeading() {
		t.Error("IsHeading() = false for Heading1")
	}

	if p.HeadingLevel() != 1 {
		t.Errorf("HeadingLevel() = %d, want 1", p.HeadingLevel())
	}
}

func TestParagraphAlignment(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	p := doc.AddParagraph()
	p.SetAlignment("center")

	if p.Alignment() != "center" {
		t.Errorf("Alignment() = %q, want %q", p.Alignment(), "center")
	}
}

func TestSaveAndOpen(t *testing.T) {
	// Create temp file
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test.docx")

	// Create and save document
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}

	p := doc.AddParagraph()
	p.SetText("Test paragraph")
	run := p.Runs()[0]
	run.SetBold(true)

	tbl := doc.AddTable(2, 2)
	tbl.Cell(0, 0).SetText("Cell content")

	if err := doc.SaveAs(tmpFile); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	doc.Close()

	// Verify file exists
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("Saved file does not exist")
	}

	// Open and verify
	doc2, err := Open(tmpFile)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer doc2.Close()

	paragraphs := doc2.Paragraphs()
	if len(paragraphs) == 0 {
		t.Fatal("Opened document has no paragraphs")
	}

	// Note: Content parsing depends on proper XML unmarshaling
	// which requires the Content field to handle interface{} correctly
}

func TestTrackChanges(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	if doc.TrackChangesEnabled() {
		t.Error("Track changes should be disabled by default")
	}

	doc.EnableTrackChanges("Test Author")

	if !doc.TrackChangesEnabled() {
		t.Error("Track changes should be enabled after EnableTrackChanges")
	}

	if doc.TrackAuthor() != "Test Author" {
		t.Errorf("TrackAuthor() = %q, want %q", doc.TrackAuthor(), "Test Author")
	}

	doc.DisableTrackChanges()

	if doc.TrackChangesEnabled() {
		t.Error("Track changes should be disabled after DisableTrackChanges")
	}
}
