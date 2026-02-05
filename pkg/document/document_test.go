package document

import (
	"fmt"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
)

// =============================================================================
// Document Creation and Save Tests
// =============================================================================

func TestDocument_New(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	if doc.Body() == nil {
		t.Error("Body() returned nil")
	}
	h.AssertParagraphCount(doc, 0)
	h.AssertTableCount(doc, 0)
}

func TestDocument_SaveAs_Fixtures(t *testing.T) {
	for _, fixture := range CommonFixtures {
		t.Run(fixture.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(fixture.Setup)
			path := h.SaveDocument(doc, fixture.Name+".docx")
			doc.Close()

			// Verify file exists and is not empty
			doc2 := h.OpenDocument(path)
			defer doc2.Close()
			// Basic sanity check - document should open without error
		})
	}
}

func TestDocument_CoreProperties(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	props := &common.CoreProperties{
		Title:          "Doc Title",
		Creator:        "Doc Author",
		Subject:        "Doc Subject",
		Description:    "Doc Description",
		Keywords:       "one;two",
		Category:       "Category",
		Language:       "en-US",
		ContentStatus:  "Draft",
		Identifier:     "urn:example:doc",
		LastModifiedBy: "Reviewer",
		Revision:       "2",
		Version:        "1.0",
		Created:        &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-02-03T00:00:00Z"},
		Modified:       &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-02-03T01:00:00Z"},
		LastPrinted:    &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-02-03T02:00:00Z"},
	}

	if err := doc.SetCoreProperties(props); err != nil {
		t.Fatalf("SetCoreProperties() error = %v", err)
	}

	got, err := doc.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties() error = %v", err)
	}
	if got.Title != props.Title {
		t.Errorf("Title = %q, want %q", got.Title, props.Title)
	}
	if got.Creator != props.Creator {
		t.Errorf("Creator = %q, want %q", got.Creator, props.Creator)
	}
	if got.Subject != props.Subject {
		t.Errorf("Subject = %q, want %q", got.Subject, props.Subject)
	}
}

// =============================================================================
// Paragraph Tests - Parameterized
// =============================================================================

func TestParagraph_SetText(t *testing.T) {
	for _, tc := range CommonTextCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			para := doc.AddParagraph()
			para.SetText(tc.Text)

			if got := para.Text(); got != tc.Text {
				t.Errorf("Text() = %q, want %q", got, tc.Text)
			}
		})
	}
}

func TestParagraph_SetStyle(t *testing.T) {
	for _, tc := range CommonHeadingStyleCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			para := doc.AddParagraph()
			para.SetStyle(tc.StyleID)

			if got := para.Style(); got != tc.StyleID {
				t.Errorf("Style() = %q, want %q", got, tc.StyleID)
			}
			if got := para.IsHeading(); got != tc.IsHeading {
				t.Errorf("IsHeading() = %v, want %v", got, tc.IsHeading)
			}
			if got := para.HeadingLevel(); got != tc.HeadingLevel {
				t.Errorf("HeadingLevel() = %d, want %d", got, tc.HeadingLevel)
			}
		})
	}
}

func TestParagraph_SetAlignment(t *testing.T) {
	for _, tc := range CommonAlignmentCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			para := doc.AddParagraph()
			para.SetAlignment(tc.Value)

			if got := para.Alignment(); got != tc.Value {
				t.Errorf("Alignment() = %q, want %q", got, tc.Value)
			}
		})
	}
}

func TestParagraph_Spacing(t *testing.T) {
	testCases := []struct {
		name   string
		before int64
		after  int64
	}{
		{"zero", 0, 0},
		{"6pt", 120, 120},   // 6pt in twips
		{"12pt", 240, 240},  // 12pt in twips
		{"asymmetric", 240, 120},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			para := doc.AddParagraph()
			para.SetSpacingBefore(tc.before)
			para.SetSpacingAfter(tc.after)

			if got := para.SpacingBefore(); got != tc.before {
				t.Errorf("SpacingBefore() = %d, want %d", got, tc.before)
			}
			if got := para.SpacingAfter(); got != tc.after {
				t.Errorf("SpacingAfter() = %d, want %d", got, tc.after)
			}
		})
	}
}

func TestParagraph_AdvancedToggles(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	para := doc.AddParagraph()
	para.SetKeepLines(true)
	para.SetPageBreakBefore(true)
	para.SetWidowControl(true)

	if !para.KeepLines() {
		t.Error("KeepLines() should be true")
	}
	if !para.PageBreakBefore() {
		t.Error("PageBreakBefore() should be true")
	}
	if !para.WidowControl() {
		t.Error("WidowControl() should be true")
	}
}

func TestParagraph_MultipleRuns(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	para := doc.AddParagraph()
	texts := []string{"Hello ", "World", "!"}
	for _, text := range texts {
		para.AddRun().SetText(text)
	}

	runs := para.Runs()
	if len(runs) != len(texts) {
		t.Fatalf("expected %d runs, got %d", len(texts), len(runs))
	}

	want := strings.Join(texts, "")
	if got := para.Text(); got != want {
		t.Errorf("Text() = %q, want %q", got, want)
	}
}

// =============================================================================
// Run Formatting Tests - Parameterized
// =============================================================================

func TestRun_FontSize(t *testing.T) {
	for _, tc := range CommonFontSizeCases {
		t.Run(fmt.Sprintf("%.1fpt", tc.Points), func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			run := doc.AddParagraph().AddRun()
			run.SetFontSize(tc.Points)

			if got := run.FontSize(); got != tc.Points {
				t.Errorf("FontSize() = %v, want %v", got, tc.Points)
			}
		})
	}
}

func TestRun_Color(t *testing.T) {
	for _, tc := range CommonColorCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			run := doc.AddParagraph().AddRun()
			run.SetColor(tc.Input)

			if got := run.Color(); got != tc.Want {
				t.Errorf("Color() = %q, want %q", got, tc.Want)
			}
		})
	}
}

func TestRun_FormattingCombinations(t *testing.T) {
	testCases := []struct {
		name   string
		bold   bool
		italic bool
		strike bool
	}{
		{"none", false, false, false},
		{"bold", true, false, false},
		{"italic", false, true, false},
		{"strike", false, false, true},
		{"bold_italic", true, true, false},
		{"bold_strike", true, false, true},
		{"italic_strike", false, true, true},
		{"all", true, true, true},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			run := doc.AddParagraph().AddRun()
			run.SetText("Test")
			run.SetBold(tc.bold)
			run.SetItalic(tc.italic)
			run.SetStrike(tc.strike)

			if run.Bold() != tc.bold {
				t.Errorf("Bold() = %v, want %v", run.Bold(), tc.bold)
			}
			if run.Italic() != tc.italic {
				t.Errorf("Italic() = %v, want %v", run.Italic(), tc.italic)
			}
			if run.Strike() != tc.strike {
				t.Errorf("Strike() = %v, want %v", run.Strike(), tc.strike)
			}
		})
	}
}

func TestRun_AdvancedEffects(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	run := doc.AddParagraph().AddRun()
	run.SetCaps(true)
	run.SetSmallCaps(true)
	run.SetOutline(true)
	run.SetShadow(true)
	run.SetEmboss(true)
	run.SetImprint(true)

	if !run.Caps() {
		t.Error("Caps() should be true")
	}
	if !run.SmallCaps() {
		t.Error("SmallCaps() should be true")
	}
	if !run.Outline() {
		t.Error("Outline() should be true")
	}
	if !run.Shadow() {
		t.Error("Shadow() should be true")
	}
	if !run.Emboss() {
		t.Error("Emboss() should be true")
	}
	if !run.Imprint() {
		t.Error("Imprint() should be true")
	}
}

func TestRun_UnderlineStyles(t *testing.T) {
	styles := []string{"single", "double", "thick", "dotted", "dash", "wave"}

	for _, style := range styles {
		t.Run(style, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			run := doc.AddParagraph().AddRun()
			run.SetUnderlineStyle(style)

			if !run.Underline() {
				t.Error("Underline() should be true")
			}
			if got := run.UnderlineStyle(); got != style {
				t.Errorf("UnderlineStyle() = %q, want %q", got, style)
			}
		})
	}
}

func TestRun_FontNames(t *testing.T) {
	fonts := []string{"Arial", "Times New Roman", "Calibri", "Courier New", "Georgia", "Verdana"}

	for _, font := range fonts {
		t.Run(font, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			run := doc.AddParagraph().AddRun()
			run.SetFontName(font)

			if got := run.FontName(); got != font {
				t.Errorf("FontName() = %q, want %q", got, font)
			}
		})
	}
}

func TestRun_Highlight(t *testing.T) {
	colors := []string{
		"yellow", "green", "cyan", "magenta", "blue", "red",
		"darkBlue", "darkCyan", "darkGreen", "darkMagenta",
		"darkRed", "darkYellow", "darkGray", "lightGray", "black",
	}

	for _, color := range colors {
		t.Run(color, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			run := doc.AddParagraph().AddRun()
			run.SetHighlight(color)

			if got := run.Highlight(); got != color {
				t.Errorf("Highlight() = %q, want %q", got, color)
			}
		})
	}
}

func TestRun_VerticalAlign(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	// Test superscript
	run1 := doc.AddParagraph().AddRun()
	run1.SetSuperscript(true)
	if !run1.Superscript() {
		t.Error("Superscript() should be true")
	}
	if run1.Subscript() {
		t.Error("Subscript() should be false")
	}

	// Test subscript
	run2 := doc.AddParagraph().AddRun()
	run2.SetSubscript(true)
	if !run2.Subscript() {
		t.Error("Subscript() should be true")
	}
	if run2.Superscript() {
		t.Error("Superscript() should be false")
	}
}

// =============================================================================
// Table Tests - Parameterized
// =============================================================================

func TestTable_Dimensions(t *testing.T) {
	for _, tc := range CommonTableDimensionCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(nil)
			defer doc.Close()

			tbl := doc.AddTable(tc.Rows, tc.Cols)

			if got := tbl.RowCount(); got != tc.Rows {
				t.Errorf("RowCount() = %d, want %d", got, tc.Rows)
			}
			if got := tbl.ColumnCount(); got != tc.Cols {
				t.Errorf("ColumnCount() = %d, want %d", got, tc.Cols)
			}
		})
	}
}

func TestTable_CellAccess(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	tbl := doc.AddTable(3, 3)

	// Valid access
	for row := 0; row < 3; row++ {
		for col := 0; col < 3; col++ {
			if cell := tbl.Cell(row, col); cell == nil {
				t.Errorf("Cell(%d, %d) returned nil", row, col)
			}
		}
	}

	// Invalid access
	invalidCases := []struct{ row, col int }{
		{-1, 0}, {0, -1}, {3, 0}, {0, 3}, {3, 3},
	}
	for _, tc := range invalidCases {
		if cell := tbl.Cell(tc.row, tc.col); cell != nil {
			t.Errorf("Cell(%d, %d) should return nil", tc.row, tc.col)
		}
	}
}

func TestTable_CellContent(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	tbl := doc.AddTable(2, 2)
	data := [][]string{
		{"A1", "B1"},
		{"A2", "B2"},
	}

	// Set content
	for row := 0; row < 2; row++ {
		for col := 0; col < 2; col++ {
			tbl.Cell(row, col).SetText(data[row][col])
		}
	}

	// Verify content
	for row := 0; row < 2; row++ {
		for col := 0; col < 2; col++ {
			got := tbl.Cell(row, col).Text()
			want := data[row][col]
			if got != want {
				t.Errorf("Cell(%d, %d).Text() = %q, want %q", row, col, got, want)
			}
		}
	}

	// Verify FirstRowText
	firstRow := tbl.FirstRowText()
	if len(firstRow) != 2 || firstRow[0] != "A1" || firstRow[1] != "B1" {
		t.Errorf("FirstRowText() = %v, want [A1 B1]", firstRow)
	}
}

func TestTable_RowOperations(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	tbl := doc.AddTable(2, 3)

	// AddRow
	tbl.AddRow()
	if got := tbl.RowCount(); got != 3 {
		t.Errorf("after AddRow: RowCount() = %d, want 3", got)
	}

	// InsertRow
	tbl.InsertRow(1)
	if got := tbl.RowCount(); got != 4 {
		t.Errorf("after InsertRow: RowCount() = %d, want 4", got)
	}

	// DeleteRow
	if err := tbl.DeleteRow(1); err != nil {
		t.Errorf("DeleteRow(1) error = %v", err)
	}
	if got := tbl.RowCount(); got != 3 {
		t.Errorf("after DeleteRow: RowCount() = %d, want 3", got)
	}

	// DeleteRow out of bounds
	if err := tbl.DeleteRow(10); err == nil {
		t.Error("DeleteRow(10) should error")
	}
}

func TestTable_CellMerge(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	tbl := doc.AddTable(3, 4)

	// Horizontal merge (gridSpan)
	tbl.Cell(0, 0).SetGridSpan(3)
	if got := tbl.Cell(0, 0).GridSpan(); got != 3 {
		t.Errorf("GridSpan() = %d, want 3", got)
	}

	// Vertical merge
	tbl.Cell(1, 0).SetVerticalMerge(VerticalMerge("restart"))
	tbl.Cell(2, 0).SetVerticalMerge(VerticalMerge("continue"))
	if got := tbl.Cell(1, 0).VerticalMerge(); got != VerticalMerge("restart") {
		t.Errorf("VerticalMerge() = %q, want restart", got)
	}
	if got := tbl.Cell(2, 0).VerticalMerge(); got != VerticalMerge("continue") {
		t.Errorf("VerticalMerge() = %q, want continue", got)
	}
}

func TestTable_Style(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	tbl := doc.AddTable(2, 2)

	if tbl.Style() != "" {
		t.Errorf("default Style() = %q, want empty", tbl.Style())
	}

	tbl.SetStyle("TableGrid")
	if got := tbl.Style(); got != "TableGrid" {
		t.Errorf("Style() = %q, want TableGrid", got)
	}
}

func TestRow_Header(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
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
}

func TestCell_Shading(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	tbl := doc.AddTable(2, 2)
	cell := tbl.Cell(0, 0)

	cell.SetShading("FFFF00")
	if got := cell.Shading(); got != "FFFF00" {
		t.Errorf("Shading() = %q, want FFFF00", got)
	}
}

// =============================================================================
// Round-Trip Tests - Using Fixtures
// =============================================================================

func TestRoundTrip_Fixtures(t *testing.T) {
	for _, fixture := range CommonFixtures {
		t.Run(fixture.Name, func(t *testing.T) {
			h := NewTestHelper(t)

			// Create, save, and reopen
			doc := h.RoundTrip(fixture.Name+".docx", fixture.Setup)
			defer doc.Close()

			// Document should open without error (basic sanity)
			if doc.Body() == nil {
				t.Error("Body() returned nil after round-trip")
			}
		})
	}
}

func TestRoundTrip_FormattingPreserved(t *testing.T) {
	h := NewTestHelper(t)

	doc := h.RoundTrip("formatting.docx", func(d Document) {
		p := d.AddParagraph()

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
	})
	defer doc.Close()

	paras := doc.Paragraphs()
	if len(paras) == 0 {
		t.Fatal("no paragraphs after round-trip")
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

func TestRoundTrip_TableContent(t *testing.T) {
	h := NewTestHelper(t)

	data := [][]string{
		{"Header1", "Header2", "Header3"},
		{"A1", "B1", "C1"},
		{"A2", "B2", "C2"},
	}

	doc := h.RoundTrip("table.docx", func(d Document) {
		tbl := d.AddTable(3, 3)
		for row := 0; row < 3; row++ {
			for col := 0; col < 3; col++ {
				tbl.Cell(row, col).SetText(data[row][col])
			}
		}
	})
	defer doc.Close()

	tables := doc.Tables()
	if len(tables) != 1 {
		t.Fatalf("expected 1 table, got %d", len(tables))
	}

	tbl := tables[0]
	for row := 0; row < 3; row++ {
		for col := 0; col < 3; col++ {
			got := tbl.Cell(row, col).Text()
			want := data[row][col]
			if got != want {
				t.Errorf("Cell(%d, %d) = %q, want %q", row, col, got, want)
			}
		}
	}
}

// =============================================================================
// Track Changes Tests
// =============================================================================

func TestTrackChanges(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	// Default state
	if doc.TrackChangesEnabled() {
		t.Error("track changes should be disabled by default")
	}

	// Enable
	doc.EnableTrackChanges("Test Author")
	if !doc.TrackChangesEnabled() {
		t.Error("track changes should be enabled")
	}
	if got := doc.TrackAuthor(); got != "Test Author" {
		t.Errorf("TrackAuthor() = %q, want 'Test Author'", got)
	}

	// Change author
	doc.SetTrackAuthor("New Author")
	if got := doc.TrackAuthor(); got != "New Author" {
		t.Errorf("TrackAuthor() = %q, want 'New Author'", got)
	}

	// Disable
	doc.DisableTrackChanges()
	if doc.TrackChangesEnabled() {
		t.Error("track changes should be disabled")
	}
}

// =============================================================================
// Body Operations Tests
// =============================================================================

func TestBody_ElementOperations(t *testing.T) {
	h := NewTestHelper(t)
	doc := h.CreateDocument(nil)
	defer doc.Close()

	body := doc.Body()

	// Initial state
	if body.ElementCount() != 0 {
		t.Errorf("initial ElementCount() = %d, want 0", body.ElementCount())
	}

	// Add elements
	body.AddParagraph().SetText("First")
	if body.ElementCount() != 1 {
		t.Errorf("after AddParagraph: ElementCount() = %d, want 1", body.ElementCount())
	}

	body.AddParagraph().SetText("Third")
	if body.ElementCount() != 2 {
		t.Errorf("after second AddParagraph: ElementCount() = %d, want 2", body.ElementCount())
	}

	// Insert at position
	body.InsertParagraphAt(1).SetText("Second")
	if body.ElementCount() != 3 {
		t.Errorf("after InsertParagraphAt: ElementCount() = %d, want 3", body.ElementCount())
	}

	// Verify order
	paras := doc.Paragraphs()
	expected := []string{"First", "Second", "Third"}
	for i, want := range expected {
		if got := paras[i].Text(); got != want {
			t.Errorf("paragraph[%d] = %q, want %q", i, got, want)
		}
	}
}
