package document

import (
	"path/filepath"
	"strings"
	"testing"
)

type documentFixtureCase struct {
	name   string
	mutate func(t *testing.T, doc Document)
	verify func(t *testing.T, doc Document)
}

func TestFixtureRoundTrip_Complex(t *testing.T) {
	cases := []documentFixtureCase{
		{
			name: "minimal.docx",
			mutate: func(t *testing.T, doc Document) {
				doc.EnableTrackChanges("Fixture Author")
				p := doc.AddParagraph()
				p.SetText("Fixture Minimal Paragraph")
				p.AddRun().SetBold(true)
				p.InsertTrackedText("Tracked")

				cc := doc.AddBlockContentControl("fixture-min", "Fixture", "Fixture CC")
				cc.SetContentControlID(42)
				_ = cc.SetContentControlLock("content")

				doc.AddHeader(HeaderFooterDefault).SetText("Fixture Header")
				doc.AddFooter(HeaderFooterDefault).SetText("Fixture Footer")

				table := doc.AddTable(2, 2)
				table.Cell(0, 0).SetText("A1")
				table.Cell(1, 1).SetText("B2")

				anchor := doc.AddParagraph()
				anchor.SetText("Fixture Anchor")
				if _, err := doc.Comments().Add("Fixture comment", "Tester", "Fixture Anchor"); err != nil {
					t.Fatalf("Add comment error = %v", err)
				}

				fieldPara := doc.AddParagraph()
				if _, err := fieldPara.AddField("DATE", "2024-01-01"); err != nil {
					t.Fatalf("AddField() error = %v", err)
				}
			},
			verify: func(t *testing.T, doc Document) {
				if doc.Header(HeaderFooterDefault) == nil || doc.Header(HeaderFooterDefault).Text() != "Fixture Header" {
					t.Error("Header text missing after round-trip")
				}
				if doc.Footer(HeaderFooterDefault) == nil || doc.Footer(HeaderFooterDefault).Text() != "Fixture Footer" {
					t.Error("Footer text missing after round-trip")
				}
				if doc.ContentControlByTag("fixture-min") == nil {
					t.Error("Content control missing after round-trip")
				}
				if !hasCommentText(doc, "Fixture comment") {
					t.Error("Expected comment text after round-trip")
				}
				if !doc.TrackChangesEnabled() || len(doc.AllRevisions()) == 0 {
					t.Error("Expected tracked revisions after round-trip")
				}
				if findParagraphContainingText(doc, "Fixture Minimal Paragraph") == nil {
					t.Error("Expected paragraph after round-trip")
				}
				if !tableHasCellText(doc, "A1") {
					t.Error("Expected table cell after round-trip")
				}
			},
		},
		{
			name: "single_paragraph.docx",
			mutate: func(t *testing.T, doc Document) {
				p := doc.AddParagraph()
				p.SetText("Fixture Single")
				if _, err := p.AddHyperlink("https://example.com", "Fixture Link"); err != nil {
					t.Fatalf("AddHyperlink() error = %v", err)
				}
				if err := p.AddBookmark("FixtureBookmark", 0, 0); err != nil {
					t.Fatalf("AddBookmark() error = %v", err)
				}
				if _, err := doc.Comments().Add("Fixture single comment", "Tester", "Fixture Single"); err != nil {
					t.Fatalf("Add comment error = %v", err)
				}
			},
			verify: func(t *testing.T, doc Document) {
				p := findParagraphContainingText(doc, "Fixture Single")
				if p == nil {
					t.Fatal("Expected paragraph after round-trip")
				}
				if len(p.Hyperlinks()) == 0 || p.Hyperlinks()[0].URL() != "https://example.com" {
					t.Error("Expected hyperlink after round-trip")
				}
				if !hasCommentText(doc, "Fixture single comment") {
					t.Error("Expected comment after round-trip")
				}
			},
		},
		{
			name: "formatted_text.docx",
			mutate: func(t *testing.T, doc Document) {
				p := doc.AddParagraph()
				p.SetAlignment("center")
				p.SetSpacingBefore(240)
				p.SetSpacingAfter(120)
				r := p.AddRun()
				r.SetText("Fixture Format")
				r.SetBold(true)
				r.SetItalic(true)
				r.SetUnderline(true)
				r.SetFontSize(14)
				r.SetFontName("Arial")
				r.SetHighlight("yellow")
				r.SetColor("FF0000")
			},
			verify: func(t *testing.T, doc Document) {
				p := findParagraphByText(doc, "Fixture Format")
				if p == nil || len(p.Runs()) == 0 {
					t.Fatal("Expected formatted run after round-trip")
				}
				if p.Alignment() != "center" {
					t.Errorf("Alignment() = %q, want center", p.Alignment())
				}
				if p.SpacingBefore() != 240 || p.SpacingAfter() != 120 {
					t.Errorf("Spacing = (%d, %d), want (240, 120)", p.SpacingBefore(), p.SpacingAfter())
				}
				r := p.Runs()[0]
				if !r.Bold() || !r.Italic() || !r.Underline() {
					t.Error("Expected run formatting after round-trip")
				}
				if r.FontSize() != 14 {
					t.Errorf("FontSize() = %v, want 14", r.FontSize())
				}
				if r.FontName() != "Arial" {
					t.Errorf("FontName() = %q, want Arial", r.FontName())
				}
				if r.Color() != "FF0000" {
					t.Errorf("Color() = %q, want FF0000", r.Color())
				}
				if r.Highlight() != "yellow" {
					t.Errorf("Highlight() = %q, want yellow", r.Highlight())
				}
			},
		},
		{
			name: "headings.docx",
			mutate: func(t *testing.T, doc Document) {
				h := doc.AddParagraph()
				h.SetText("Fixture Heading")
				h.SetStyle("Heading2")
			},
			verify: func(t *testing.T, doc Document) {
				p := findParagraphByText(doc, "Fixture Heading")
				if p == nil || p.Style() != "Heading2" {
					t.Error("Expected heading style after round-trip")
				}
			},
		},
		{
			name: "simple_table.docx",
			mutate: func(t *testing.T, doc Document) {
				tables := doc.Tables()
				if len(tables) == 0 {
					t.Fatal("Expected fixture table")
				}
				row := tables[0].AddRow()
				row.Cell(0).SetText("Fixture Row")
			},
			verify: func(t *testing.T, doc Document) {
				if !tableHasCellText(doc, "Fixture Row") {
					t.Error("Expected table row after round-trip")
				}
			},
		},
		{
			name: "complex_table.docx",
			mutate: func(t *testing.T, doc Document) {
				table := doc.AddTable(1, 1)
				cell := table.Cell(0, 0)
				cell.SetText("Fixture Complex")
				cell.SetShading("CCCCCC")
			},
			verify: func(t *testing.T, doc Document) {
				if !tableHasCellText(doc, "Fixture Complex") {
					t.Error("Expected appended table after round-trip")
				}
			},
		},
		{
			name: "track_changes.docx",
			mutate: func(t *testing.T, doc Document) {
				doc.EnableTrackChanges("Fixture Author")
				p := doc.AddParagraph()
				p.InsertTrackedText("Fixture Tracked")
			},
			verify: func(t *testing.T, doc Document) {
				if !doc.TrackChangesEnabled() || len(doc.AllRevisions()) == 0 {
					t.Error("Expected tracked revisions after round-trip")
				}
			},
		},
		{
			name: "comments.docx",
			mutate: func(t *testing.T, doc Document) {
				doc.AddParagraph().SetText("Comment Target")
				comment, err := doc.Comments().Add("Fixture comment", "Tester", "Comment Target")
				if err != nil {
					t.Fatalf("Add comment error = %v", err)
				}
				if _, err := comment.AddReply("Fixture reply", "Reviewer"); err != nil {
					t.Fatalf("AddReply error = %v", err)
				}
			},
			verify: func(t *testing.T, doc Document) {
				if !hasCommentText(doc, "Fixture comment") {
					t.Error("Expected comment after round-trip")
				}
				foundReply := false
				for _, comment := range doc.Comments().All() {
					for _, reply := range comment.Replies() {
						if reply.Text() == "Fixture reply" {
							foundReply = true
							break
						}
					}
				}
				if !foundReply {
					t.Error("Expected comment reply after round-trip")
				}
			},
		},
		{
			name: "styles.docx",
			mutate: func(t *testing.T, doc Document) {
				style := doc.AddParagraphStyle("FixtureStyle", "Fixture Style")
				style.SetBold(true)
				p := doc.AddParagraph()
				p.SetText("Styled Text")
				p.SetStyle("FixtureStyle")
			},
			verify: func(t *testing.T, doc Document) {
				if doc.StyleByID("FixtureStyle") == nil {
					t.Error("Expected style after round-trip")
				}
				p := findParagraphByText(doc, "Styled Text")
				if p == nil || p.Style() != "FixtureStyle" {
					t.Error("Expected styled paragraph after round-trip")
				}
			},
		},
		{
			name: "headers_footers.docx",
			mutate: func(t *testing.T, doc Document) {
				header := doc.Header(HeaderFooterDefault)
				if header == nil {
					header = doc.AddHeader(HeaderFooterDefault)
				}
				header.SetText("Fixture Header")
				footer := doc.Footer(HeaderFooterDefault)
				if footer == nil {
					footer = doc.AddFooter(HeaderFooterDefault)
				}
				footer.SetText("Fixture Footer")
				doc.AddParagraph().SetText("Fixture HF")
				if sections := doc.Sections(); len(sections) > 0 {
					sections[0].SetPageMargins(PageMargins{
						Top: 1440, Bottom: 1440, Left: 1800, Right: 1800,
						Header: 720, Footer: 720, Gutter: 0,
					})
				}
			},
			verify: func(t *testing.T, doc Document) {
				if doc.Header(HeaderFooterDefault) == nil || doc.Header(HeaderFooterDefault).Text() != "Fixture Header" {
					t.Error("Header text missing after round-trip")
				}
				if doc.Footer(HeaderFooterDefault) == nil || doc.Footer(HeaderFooterDefault).Text() != "Fixture Footer" {
					t.Error("Footer text missing after round-trip")
				}
				if sections := doc.Sections(); len(sections) > 0 {
					if margins, ok := sections[0].PageMargins(); !ok || margins.Left != 1800 || margins.Right != 1800 {
						t.Error("Expected page margins after round-trip")
					}
				}
			},
		},
		{
			name: "sdt_content_controls.docx",
			mutate: func(t *testing.T, doc Document) {
				cc := doc.AddBlockContentControl("FixtureTag", "Fixture", "Fixture CC")
				cc.SetDateConfig(ContentControlDateConfig{Format: "yyyy-MM-dd"})
			},
			verify: func(t *testing.T, doc Document) {
				cc := doc.ContentControlByTag("FixtureTag")
				if cc == nil || !strings.Contains(cc.Text(), "Fixture CC") {
					t.Error("Expected content control after round-trip")
				}
			},
		},
		{
			name: "numbered_list.docx",
			mutate: func(t *testing.T, doc Document) {
				numID, err := doc.AddNumberedListStyle()
				if err != nil {
					t.Fatalf("AddNumberedListStyle() error = %v", err)
				}
				p := doc.AddParagraph()
				p.SetText("Fixture Numbered")
				if err := p.SetList(numID, 0); err != nil {
					t.Fatalf("SetList() error = %v", err)
				}
			},
			verify: func(t *testing.T, doc Document) {
				p := findParagraphByText(doc, "Fixture Numbered")
				if p == nil || p.ListLevel() != 0 {
					t.Error("Expected numbered list after round-trip")
				}
			},
		},
		{
			name: "bullet_list.docx",
			mutate: func(t *testing.T, doc Document) {
				p := doc.AddParagraph()
				p.SetText("Fixture Bullet")
				if err := p.SetListLevel(0); err != nil {
					t.Fatalf("SetListLevel() error = %v", err)
				}
			},
			verify: func(t *testing.T, doc Document) {
				p := findParagraphByText(doc, "Fixture Bullet")
				if p == nil || p.ListLevel() != 0 {
					t.Error("Expected list after round-trip")
				}
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.OpenFixture(tc.name)
			if tc.mutate != nil {
				tc.mutate(t, doc)
			}

			outPath := filepath.Join(t.TempDir(), tc.name)
			if err := doc.SaveAs(outPath); err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			if err := doc.Close(); err != nil {
				t.Fatalf("Close() error = %v", err)
			}

			round, err := Open(outPath)
			if err != nil {
				t.Fatalf("Open(roundtrip) error = %v", err)
			}
			defer round.Close()

			if tc.verify != nil {
				tc.verify(t, round)
			}
		})
	}
}

func findParagraphByText(doc Document, text string) Paragraph {
	for _, para := range doc.Paragraphs() {
		if para.Text() == text {
			return para
		}
	}
	return nil
}

func findParagraphContainingText(doc Document, text string) Paragraph {
	for _, para := range doc.Paragraphs() {
		if strings.Contains(para.Text(), text) {
			return para
		}
	}
	return nil
}

func hasCommentText(doc Document, text string) bool {
	for _, comment := range doc.Comments().All() {
		if comment.Text() == text {
			return true
		}
	}
	return false
}

func tableHasCellText(doc Document, text string) bool {
	for _, table := range doc.Tables() {
		for r := 0; r < table.RowCount(); r++ {
			for c := 0; c < table.ColumnCount(); c++ {
				cell := table.Cell(r, c)
				if cell != nil && cell.Text() == text {
					return true
				}
			}
		}
	}
	return false
}
