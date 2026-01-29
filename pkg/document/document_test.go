package document

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// =============================================================================
// Document Tests
// =============================================================================

func TestNewDocument(t *testing.T) {
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	if doc.Body() == nil {
		t.Error("Body() returned nil")
	}

	if len(doc.Paragraphs()) != 0 {
		t.Errorf("New document has %d paragraphs, want 0", len(doc.Paragraphs()))
	}

	if len(doc.Tables()) != 0 {
		t.Errorf("New document has %d tables, want 0", len(doc.Tables()))
	}
}

func TestDocument_SaveAs(t *testing.T) {
	tmpDir := t.TempDir()

	tests := []struct {
		name     string
		setup    func(*Document)
		filename string
	}{
		{
			name:     "empty document",
			setup:    func(d *Document) {},
			filename: "empty.docx",
		},
		{
			name: "with paragraph",
			setup: func(d *Document) {
				d.AddParagraph().SetText("Hello")
			},
			filename: "with_para.docx",
		},
		{
			name: "with table",
			setup: func(d *Document) {
				d.AddTable(2, 2)
			},
			filename: "with_table.docx",
		},
		{
			name: "complex document",
			setup: func(d *Document) {
				p := d.AddParagraph()
				p.SetText("Title")
				p.SetStyle("Heading1")
				d.AddParagraph().SetText("Body text")
				tbl := d.AddTable(3, 3)
				tbl.Cell(0, 0).SetText("Header")
			},
			filename: "complex.docx",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, err := New()
			if err != nil {
				t.Fatalf("New() error = %v", err)
			}

			tt.setup(doc)

			path := filepath.Join(tmpDir, tt.filename)
			err = doc.SaveAs(path)
			if err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			doc.Close()

			// Verify file exists and is not empty
			info, err := os.Stat(path)
			if err != nil {
				t.Fatalf("file not created: %v", err)
			}
			if info.Size() == 0 {
				t.Error("file is empty")
			}
		})
	}
}

// =============================================================================
// Paragraph Tests - Table Driven
// =============================================================================

func TestParagraph_SetStyle(t *testing.T) {
	tests := []struct {
		name      string
		styleID   string
		wantLevel int
	}{
		{"Heading1", "Heading1", 1},
		{"Heading2", "Heading2", 2},
		{"Heading3", "Heading3", 3},
		{"Heading9", "Heading9", 9},
		{"Normal", "Normal", 0},
		{"Empty style", "", 0},
		{"Custom style", "CustomHeading", 0},
		{"Title", "Title", 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			para := doc.AddParagraph()

			para.SetStyle(tt.styleID)

			if got := para.Style(); got != tt.styleID {
				t.Errorf("Style() = %v, want %v", got, tt.styleID)
			}
			if got := para.HeadingLevel(); got != tt.wantLevel {
				t.Errorf("HeadingLevel() = %v, want %v", got, tt.wantLevel)
			}
		})
	}
}

func TestParagraph_IsHeading(t *testing.T) {
	tests := []struct {
		style string
		want  bool
	}{
		{"Heading1", true},
		{"Heading2", true},
		{"heading1", true}, // case variations
		{"Normal", false},
		{"Title", false},
		{"", false},
	}

	for _, tt := range tests {
		t.Run(tt.style, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			para := doc.AddParagraph()
			para.SetStyle(tt.style)

			if got := para.IsHeading(); got != tt.want {
				t.Errorf("IsHeading() = %v, want %v for style %q", got, tt.want, tt.style)
			}
		})
	}
}

func TestParagraph_Alignment(t *testing.T) {
	tests := []struct {
		alignment string
	}{
		{"left"},
		{"center"},
		{"right"},
		{"both"}, // justified
		{""},
	}

	for _, tt := range tests {
		t.Run(tt.alignment, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			para := doc.AddParagraph()

			if tt.alignment != "" {
				para.SetAlignment(tt.alignment)
			}

			if got := para.Alignment(); got != tt.alignment {
				t.Errorf("Alignment() = %q, want %q", got, tt.alignment)
			}
		})
	}
}

func TestParagraph_Spacing(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	para := doc.AddParagraph()

	// Test default values
	if para.SpacingBefore() != 0 {
		t.Errorf("default SpacingBefore() = %d, want 0", para.SpacingBefore())
	}
	if para.SpacingAfter() != 0 {
		t.Errorf("default SpacingAfter() = %d, want 0", para.SpacingAfter())
	}

	// Set and verify
	para.SetSpacingBefore(240) // 12pt in twips
	para.SetSpacingAfter(120)  // 6pt in twips

	if para.SpacingBefore() != 240 {
		t.Errorf("SpacingBefore() = %d, want 240", para.SpacingBefore())
	}
	if para.SpacingAfter() != 120 {
		t.Errorf("SpacingAfter() = %d, want 120", para.SpacingAfter())
	}
}

func TestParagraph_KeepWithNext(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	para := doc.AddParagraph()

	if para.KeepWithNext() {
		t.Error("default KeepWithNext() should be false")
	}

	para.SetKeepWithNext(true)
	if !para.KeepWithNext() {
		t.Error("KeepWithNext() should be true after SetKeepWithNext(true)")
	}

	para.SetKeepWithNext(false)
	if para.KeepWithNext() {
		t.Error("KeepWithNext() should be false after SetKeepWithNext(false)")
	}
}

func TestParagraph_Text(t *testing.T) {
	tests := []struct {
		name string
		text string
	}{
		{"simple", "Hello World"},
		{"empty", ""},
		{"unicode", "Hello 世界 🌍"},
		{"whitespace", "  spaces  "},
		{"multiword", "The quick brown fox"},
		{"special chars", "Line1\tTab\nNewline"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			para := doc.AddParagraph()

			para.SetText(tt.text)

			if got := para.Text(); got != tt.text {
				t.Errorf("Text() = %q, want %q", got, tt.text)
			}
		})
	}
}

func TestParagraph_MultipleRuns(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	para := doc.AddParagraph()

	run1 := para.AddRun()
	run1.SetText("Hello ")
	run1.SetBold(true)

	run2 := para.AddRun()
	run2.SetText("World")
	run2.SetItalic(true)

	runs := para.Runs()
	if len(runs) != 2 {
		t.Fatalf("expected 2 runs, got %d", len(runs))
	}

	if para.Text() != "Hello World" {
		t.Errorf("Text() = %q, want %q", para.Text(), "Hello World")
	}
}

// =============================================================================
// Run Tests - Table Driven
// =============================================================================

func TestRun_TextOperations(t *testing.T) {
	tests := []struct {
		name string
		text string
	}{
		{"simple", "Hello"},
		{"empty", ""},
		{"unicode", "日本語"},
		{"emoji", "🎉🎊"},
		{"long text", strings.Repeat("abc", 100)},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()

			run.SetText(tt.text)

			if got := run.Text(); got != tt.text {
				t.Errorf("Text() = %q, want %q", got, tt.text)
			}
		})
	}
}

func TestRun_BoldItalicStrike(t *testing.T) {
	tests := []struct {
		name   string
		bold   bool
		italic bool
		strike bool
	}{
		{"none", false, false, false},
		{"bold only", true, false, false},
		{"italic only", false, true, false},
		{"strike only", false, false, true},
		{"bold+italic", true, true, false},
		{"all", true, true, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()
			run.SetText("Test")

			run.SetBold(tt.bold)
			run.SetItalic(tt.italic)
			run.SetStrike(tt.strike)

			if run.Bold() != tt.bold {
				t.Errorf("Bold() = %v, want %v", run.Bold(), tt.bold)
			}
			if run.Italic() != tt.italic {
				t.Errorf("Italic() = %v, want %v", run.Italic(), tt.italic)
			}
			if run.Strike() != tt.strike {
				t.Errorf("Strike() = %v, want %v", run.Strike(), tt.strike)
			}
		})
	}
}

func TestRun_Underline(t *testing.T) {
	tests := []struct {
		name      string
		style     string
		underline bool
	}{
		{"none", "", false},
		{"single", "single", true},
		{"double", "double", true},
		{"wave", "wave", true},
		{"dotted", "dotted", true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()

			if tt.style != "" {
				run.SetUnderlineStyle(tt.style)
			}

			if run.Underline() != tt.underline {
				t.Errorf("Underline() = %v, want %v", run.Underline(), tt.underline)
			}
			if tt.style != "" && run.UnderlineStyle() != tt.style {
				t.Errorf("UnderlineStyle() = %q, want %q", run.UnderlineStyle(), tt.style)
			}
		})
	}
}

func TestRun_FontSize(t *testing.T) {
	tests := []struct {
		name   string
		points float64
	}{
		{"8pt", 8},
		{"10pt", 10},
		{"12pt", 12},
		{"14pt", 14},
		{"18pt", 18},
		{"24pt", 24},
		{"36pt", 36},
		{"72pt", 72},
		{"10.5pt", 10.5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()

			run.SetFontSize(tt.points)

			if got := run.FontSize(); got != tt.points {
				t.Errorf("FontSize() = %v, want %v", got, tt.points)
			}
		})
	}
}

func TestRun_FontName(t *testing.T) {
	tests := []struct {
		font string
	}{
		{"Arial"},
		{"Times New Roman"},
		{"Calibri"},
		{"Courier New"},
		{"Georgia"},
		{"Verdana"},
	}

	for _, tt := range tests {
		t.Run(tt.font, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()

			run.SetFontName(tt.font)

			if got := run.FontName(); got != tt.font {
				t.Errorf("FontName() = %q, want %q", got, tt.font)
			}
		})
	}
}

func TestRun_Color(t *testing.T) {
	tests := []struct {
		name  string
		color string
	}{
		{"red", "FF0000"},
		{"green", "00FF00"},
		{"blue", "0000FF"},
		{"black", "000000"},
		{"white", "FFFFFF"},
		{"with hash", "#FF0000"}, // Should strip #
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()

			run.SetColor(tt.color)

			want := strings.TrimPrefix(tt.color, "#")
			if got := run.Color(); got != want {
				t.Errorf("Color() = %q, want %q", got, want)
			}
		})
	}
}

func TestRun_Highlight(t *testing.T) {
	tests := []struct {
		color string
	}{
		{"yellow"},
		{"green"},
		{"cyan"},
		{"magenta"},
		{"blue"},
		{"red"},
		{"darkBlue"},
		{"darkCyan"},
		{"darkGreen"},
		{"darkMagenta"},
		{"darkRed"},
		{"darkYellow"},
		{"darkGray"},
		{"lightGray"},
		{"black"},
	}

	for _, tt := range tests {
		t.Run(tt.color, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			run := doc.AddParagraph().AddRun()

			run.SetHighlight(tt.color)

			if got := run.Highlight(); got != tt.color {
				t.Errorf("Highlight() = %q, want %q", got, tt.color)
			}
		})
	}
}

func TestRun_SuperscriptSubscript(t *testing.T) {
	doc, _ := New()
	defer doc.Close()

	// Test superscript
	run1 := doc.AddParagraph().AddRun()
	run1.SetSuperscript(true)
	if !run1.Superscript() {
		t.Error("Superscript() should be true")
	}
	if run1.Subscript() {
		t.Error("Subscript() should be false when superscript")
	}

	// Test subscript
	run2 := doc.AddParagraph().AddRun()
	run2.SetSubscript(true)
	if !run2.Subscript() {
		t.Error("Subscript() should be true")
	}
	if run2.Superscript() {
		t.Error("Superscript() should be false when subscript")
	}
}

func TestRun_Breaks(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	run := doc.AddParagraph().AddRun()

	run.SetText("Before")
	run.AddBreak()
	// Note: AddBreak adds to content but Text() only reads text elements

	run.AddTab()
	run.AddPageBreak()

	// Run should not error
	_ = run.Text()
}

// =============================================================================
// Table Tests - Table Driven
// =============================================================================

func TestTable_Dimensions(t *testing.T) {
	tests := []struct {
		name string
		rows int
		cols int
	}{
		{"1x1", 1, 1},
		{"2x2", 2, 2},
		{"3x3", 3, 3},
		{"5x3", 5, 3},
		{"3x5", 3, 5},
		{"10x10", 10, 10},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc, _ := New()
			defer doc.Close()
			tbl := doc.AddTable(tt.rows, tt.cols)

			if tbl.RowCount() != tt.rows {
				t.Errorf("RowCount() = %d, want %d", tbl.RowCount(), tt.rows)
			}
			if tbl.ColumnCount() != tt.cols {
				t.Errorf("ColumnCount() = %d, want %d", tbl.ColumnCount(), tt.cols)
			}
		})
	}
}

func TestTable_CellAccess(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(3, 3)

	// Test all cells are accessible
	for row := 0; row < 3; row++ {
		for col := 0; col < 3; col++ {
			cell := tbl.Cell(row, col)
			if cell == nil {
				t.Errorf("Cell(%d, %d) returned nil", row, col)
			}
		}
	}

	// Test out of bounds
	if tbl.Cell(-1, 0) != nil {
		t.Error("Cell(-1, 0) should return nil")
	}
	if tbl.Cell(0, -1) != nil {
		t.Error("Cell(0, -1) should return nil")
	}
	if tbl.Cell(3, 0) != nil {
		t.Error("Cell(3, 0) should return nil for 3-row table")
	}
	if tbl.Cell(0, 3) != nil {
		t.Error("Cell(0, 3) should return nil for 3-col table")
	}
}

func TestTable_CellText(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 2)

	data := [][]string{
		{"A1", "B1"},
		{"A2", "B2"},
	}

	// Set cell text
	for row := 0; row < 2; row++ {
		for col := 0; col < 2; col++ {
			tbl.Cell(row, col).SetText(data[row][col])
		}
	}

	// Verify cell text
	for row := 0; row < 2; row++ {
		for col := 0; col < 2; col++ {
			got := tbl.Cell(row, col).Text()
			want := data[row][col]
			if got != want {
				t.Errorf("Cell(%d, %d).Text() = %q, want %q", row, col, got, want)
			}
		}
	}
}

func TestTable_AddRow(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 3)

	initialRows := tbl.RowCount()
	newRow := tbl.AddRow()

	if tbl.RowCount() != initialRows+1 {
		t.Errorf("RowCount() = %d after AddRow, want %d", tbl.RowCount(), initialRows+1)
	}

	// New row should have correct number of cells
	if len(newRow.Cells()) != 3 {
		t.Errorf("new row has %d cells, want 3", len(newRow.Cells()))
	}
}

func TestTable_InsertRow(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(3, 2)

	// Set initial data
	tbl.Cell(0, 0).SetText("Row0")
	tbl.Cell(1, 0).SetText("Row1")
	tbl.Cell(2, 0).SetText("Row2")

	// Insert at middle
	newRow := tbl.InsertRow(1)
	newRow.Cell(0).SetText("Inserted")

	if tbl.RowCount() != 4 {
		t.Errorf("RowCount() = %d after InsertRow, want 4", tbl.RowCount())
	}

	// Verify order
	expected := []string{"Row0", "Inserted", "Row1", "Row2"}
	for i, want := range expected {
		got := tbl.Cell(i, 0).Text()
		if got != want {
			t.Errorf("row %d text = %q, want %q", i, got, want)
		}
	}
}

func TestTable_DeleteRow(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(3, 2)

	tbl.Cell(0, 0).SetText("Row0")
	tbl.Cell(1, 0).SetText("Row1")
	tbl.Cell(2, 0).SetText("Row2")

	err := tbl.DeleteRow(1)
	if err != nil {
		t.Fatalf("DeleteRow(1) error = %v", err)
	}

	if tbl.RowCount() != 2 {
		t.Errorf("RowCount() = %d after DeleteRow, want 2", tbl.RowCount())
	}

	// Verify remaining rows
	if tbl.Cell(0, 0).Text() != "Row0" {
		t.Error("first row incorrect after delete")
	}
	if tbl.Cell(1, 0).Text() != "Row2" {
		t.Error("second row incorrect after delete")
	}

	// Test invalid index
	err = tbl.DeleteRow(10)
	if err == nil {
		t.Error("DeleteRow(10) should error for out of bounds")
	}
}

func TestTable_FirstRowText(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 3)

	tbl.Cell(0, 0).SetText("Header1")
	tbl.Cell(0, 1).SetText("Header2")
	tbl.Cell(0, 2).SetText("Header3")

	headers := tbl.FirstRowText()
	if len(headers) != 3 {
		t.Fatalf("FirstRowText() returned %d items, want 3", len(headers))
	}

	expected := []string{"Header1", "Header2", "Header3"}
	for i, want := range expected {
		if headers[i] != want {
			t.Errorf("header[%d] = %q, want %q", i, headers[i], want)
		}
	}
}

func TestTable_Style(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 2)

	if tbl.Style() != "" {
		t.Errorf("default Style() = %q, want empty", tbl.Style())
	}

	tbl.SetStyle("TableGrid")
	if tbl.Style() != "TableGrid" {
		t.Errorf("Style() = %q, want TableGrid", tbl.Style())
	}
}

func TestRow_HeaderRow(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(3, 2)

	row := tbl.Row(0)
	if row.IsHeader() {
		t.Error("default IsHeader() should be false")
	}

	row.SetHeader(true)
	if !row.IsHeader() {
		t.Error("IsHeader() should be true after SetHeader(true)")
	}

	row.SetHeader(false)
	if row.IsHeader() {
		t.Error("IsHeader() should be false after SetHeader(false)")
	}
}

func TestCell_GridSpan(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 4)

	cell := tbl.Cell(0, 0)
	if cell.GridSpan() != 1 {
		t.Errorf("default GridSpan() = %d, want 1", cell.GridSpan())
	}

	cell.SetGridSpan(3)
	if cell.GridSpan() != 3 {
		t.Errorf("GridSpan() = %d, want 3", cell.GridSpan())
	}

	cell.SetGridSpan(1) // Reset
	if cell.GridSpan() != 1 {
		t.Errorf("GridSpan() = %d after reset, want 1", cell.GridSpan())
	}
}

func TestCell_VerticalMerge(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(3, 2)

	// First cell starts merge
	tbl.Cell(0, 0).SetVerticalMerge("restart")
	// Following cells continue
	tbl.Cell(1, 0).SetVerticalMerge("continue")
	tbl.Cell(2, 0).SetVerticalMerge("continue")

	if tbl.Cell(0, 0).VerticalMerge() != "restart" {
		t.Error("first cell should have restart merge")
	}
	if tbl.Cell(1, 0).VerticalMerge() != "continue" {
		t.Error("second cell should have continue merge")
	}
}

func TestCell_Shading(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 2)

	cell := tbl.Cell(0, 0)
	cell.SetShading("FFFF00") // Yellow

	if cell.Shading() != "FFFF00" {
		t.Errorf("Shading() = %q, want FFFF00", cell.Shading())
	}
}

func TestCell_Paragraphs(t *testing.T) {
	doc, _ := New()
	defer doc.Close()
	tbl := doc.AddTable(2, 2)

	cell := tbl.Cell(0, 0)
	cell.SetText("First paragraph")
	cell.AddParagraph().SetText("Second paragraph")

	paras := cell.Paragraphs()
	if len(paras) != 2 {
		t.Fatalf("cell has %d paragraphs, want 2", len(paras))
	}
}

// =============================================================================
// Track Changes Tests
// =============================================================================

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

	doc.SetTrackAuthor("New Author")
	if doc.TrackAuthor() != "New Author" {
		t.Errorf("TrackAuthor() = %q after SetTrackAuthor, want %q", doc.TrackAuthor(), "New Author")
	}

	doc.DisableTrackChanges()

	if doc.TrackChangesEnabled() {
		t.Error("Track changes should be disabled after DisableTrackChanges")
	}
}

// =============================================================================
// Round-Trip Tests
// =============================================================================

func TestDocument_RoundTrip_Empty(t *testing.T) {
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "roundtrip.docx")

	// Create and save
	doc, _ := New()
	doc.SaveAs(tmpFile)
	doc.Close()

	// Re-open
	doc2, err := Open(tmpFile)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer doc2.Close()

	if len(doc2.Paragraphs()) != 0 {
		t.Error("empty document should have no paragraphs after round-trip")
	}
}

func TestDocument_RoundTrip_WithContent(t *testing.T) {
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "roundtrip.docx")

	// Create document with content
	doc, _ := New()

	p1 := doc.AddParagraph()
	p1.SetText("Hello World")
	p1.SetStyle("Heading1")

	p2 := doc.AddParagraph()
	run := p2.AddRun()
	run.SetText("Formatted text")
	run.SetBold(true)
	run.SetItalic(true)

	tbl := doc.AddTable(2, 2)
	tbl.Cell(0, 0).SetText("A1")
	tbl.Cell(0, 1).SetText("B1")
	tbl.Cell(1, 0).SetText("A2")
	tbl.Cell(1, 1).SetText("B2")

	doc.SaveAs(tmpFile)
	doc.Close()

	// Re-open and verify
	doc2, err := Open(tmpFile)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer doc2.Close()

	// Check paragraph count
	paras := doc2.Paragraphs()
	if len(paras) < 2 {
		t.Errorf("expected at least 2 paragraphs, got %d", len(paras))
	}

	// Check first paragraph
	if len(paras) > 0 {
		if paras[0].Text() != "Hello World" {
			t.Errorf("first paragraph text = %q, want %q", paras[0].Text(), "Hello World")
		}
		if paras[0].Style() != "Heading1" {
			t.Errorf("first paragraph style = %q, want Heading1", paras[0].Style())
		}
	}

	// Check table count
	tables := doc2.Tables()
	if len(tables) != 1 {
		t.Errorf("expected 1 table, got %d", len(tables))
	}

	// Check table content
	if len(tables) > 0 {
		if tables[0].Cell(0, 0).Text() != "A1" {
			t.Errorf("table cell(0,0) = %q, want A1", tables[0].Cell(0, 0).Text())
		}
	}
}

func TestDocument_RoundTrip_Formatting(t *testing.T) {
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "formatting.docx")

	// Create with various formatting
	doc, _ := New()
	p := doc.AddParagraph()

	r1 := p.AddRun()
	r1.SetText("Bold ")
	r1.SetBold(true)

	r2 := p.AddRun()
	r2.SetText("Italic ")
	r2.SetItalic(true)

	r3 := p.AddRun()
	r3.SetText("Colored")
	r3.SetColor("FF0000")
	r3.SetFontSize(14)
	r3.SetFontName("Arial")

	doc.SaveAs(tmpFile)
	doc.Close()

	// Re-open and verify formatting persisted
	doc2, err := Open(tmpFile)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer doc2.Close()

	paras := doc2.Paragraphs()
	if len(paras) == 0 {
		t.Fatal("no paragraphs found")
	}

	runs := paras[0].Runs()
	if len(runs) < 3 {
		t.Fatalf("expected 3 runs, got %d", len(runs))
	}

	if !runs[0].Bold() {
		t.Error("first run should be bold")
	}
	if !runs[1].Italic() {
		t.Error("second run should be italic")
	}
	if runs[2].Color() != "FF0000" {
		t.Errorf("third run color = %q, want FF0000", runs[2].Color())
	}
	if runs[2].FontSize() != 14 {
		t.Errorf("third run font size = %v, want 14", runs[2].FontSize())
	}
	if runs[2].FontName() != "Arial" {
		t.Errorf("third run font = %q, want Arial", runs[2].FontName())
	}
}

// =============================================================================
// Body Tests
// =============================================================================

func TestBody_ElementCount(t *testing.T) {
	doc, _ := New()
	defer doc.Close()

	body := doc.Body()
	if body.ElementCount() != 0 {
		t.Errorf("new document body has %d elements, want 0", body.ElementCount())
	}

	doc.AddParagraph()
	if body.ElementCount() != 1 {
		t.Errorf("after AddParagraph, body has %d elements, want 1", body.ElementCount())
	}

	doc.AddTable(2, 2)
	if body.ElementCount() != 2 {
		t.Errorf("after AddTable, body has %d elements, want 2", body.ElementCount())
	}
}

func TestBody_InsertParagraphAt(t *testing.T) {
	doc, _ := New()
	defer doc.Close()

	body := doc.Body()

	p1 := body.AddParagraph()
	p1.SetText("First")

	p2 := body.AddParagraph()
	p2.SetText("Third")

	// Insert at position 1
	inserted := body.InsertParagraphAt(1)
	inserted.SetText("Second")

	paras := doc.Paragraphs()
	if len(paras) != 3 {
		t.Fatalf("expected 3 paragraphs, got %d", len(paras))
	}

	expected := []string{"First", "Second", "Third"}
	for i, want := range expected {
		if paras[i].Text() != want {
			t.Errorf("paragraph %d = %q, want %q", i, paras[i].Text(), want)
		}
	}
}
