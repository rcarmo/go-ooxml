package document

import (
	"path/filepath"
	"testing"
)

// ECMA-376 Validation Tests
// These tests validate our output against the official Open XML SDK validator.
// They verify compliance with ECMA-376 Part 1 (WordprocessingML) specifications.

// =============================================================================
// ECMA-376 Compliance Tests - Using Test Fixtures
// =============================================================================

// TestECMA376_Fixtures validates all common fixtures against the schema.
func TestECMA376_Fixtures(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)
	
	for _, fixture := range CommonFixtures {
		t.Run(fixture.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(fixture.Setup)
			path := h.SaveDocument(doc, fixture.Name+".docx")
			doc.Close()
			
			validator.AssertValid(t, path)
		})
	}
}

// TestECMA376_ParagraphStyles validates paragraph style references.
// Per ECMA-376 §17.3.1.27, w:pStyle references must use w:val attribute.
func TestECMA376_ParagraphStyles(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)

	for _, tc := range CommonHeadingStyleCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d *Document) {
				para := d.AddParagraph()
				para.SetText("Test " + tc.StyleID)
				para.SetStyle(tc.StyleID)
			})
			path := h.SaveDocument(doc, "style.docx")
			doc.Close()

			validator.AssertValid(t, path)
		})
	}
}

// TestECMA376_RunFormatting validates run-level formatting.
// Per ECMA-376 §17.3.2, w:rPr contains formatting like w:b, w:i, w:u, w:sz.
func TestECMA376_RunFormatting(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)
	h := NewTestHelper(t)

	doc := h.CreateDocument(func(d *Document) {
		para := d.AddParagraph()

		r1 := para.AddRun()
		r1.SetText("Bold ")
		r1.SetBold(true)

		r2 := para.AddRun()
		r2.SetText("Italic ")
		r2.SetItalic(true)

		r3 := para.AddRun()
		r3.SetText("Underline ")
		r3.SetUnderline(true)

		r4 := para.AddRun()
		r4.SetText("Large ")
		r4.SetFontSize(24)

		r5 := para.AddRun()
		r5.SetText("Red")
		r5.SetColor("FF0000")
	})
	path := h.SaveDocument(doc, "formatting.docx")
	doc.Close()

	validator.AssertValid(t, path)
}

// TestECMA376_FontSizes validates font size values.
// Per ECMA-376 §17.3.2.38, w:sz/@w:val is in half-points (1pt = 2 half-points).
func TestECMA376_FontSizes(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)

	for _, tc := range CommonFontSizeCases {
		t.Run(formatFloat(tc.Points), func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d *Document) {
				run := d.AddParagraph().AddRun()
				run.SetText("Test")
				run.SetFontSize(tc.Points)
			})
			path := h.SaveDocument(doc, "size.docx")
			doc.Close()

			validator.AssertValid(t, path)
		})
	}
}

// TestECMA376_TableDimensions validates various table sizes.
// Per ECMA-376 §17.4.38, w:tbl must contain w:tblPr, w:tblGrid, and w:tr elements.
func TestECMA376_TableDimensions(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)

	for _, tc := range CommonTableDimensionCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d *Document) {
				tbl := d.AddTable(tc.Rows, tc.Cols)
				tbl.Cell(0, 0).SetText("Test")
			})
			path := h.SaveDocument(doc, "table.docx")
			doc.Close()

			validator.AssertValid(t, path)
		})
	}
}

// TestECMA376_ParagraphAlignment validates alignment values.
// Per ECMA-376 §17.3.1.13, w:jc/@w:val must be one of: start, end, center, both, etc.
func TestECMA376_ParagraphAlignment(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)

	for _, tc := range CommonAlignmentCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d *Document) {
				para := d.AddParagraph()
				para.SetText("Aligned text")
				para.SetAlignment(tc.Value)
			})
			path := h.SaveDocument(doc, "alignment.docx")
			doc.Close()

			validator.AssertValid(t, path)
		})
	}
}

// TestECMA376_TextContent validates various text content.
func TestECMA376_TextContent(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)

	for _, tc := range CommonTextCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d *Document) {
				d.AddParagraph().SetText(tc.Text)
			})
			path := h.SaveDocument(doc, "text.docx")
			doc.Close()

			validator.AssertValid(t, path)
		})
	}
}

// TestECMA376_Colors validates color values.
func TestECMA376_Colors(t *testing.T) {
	validator := NewValidator()
	validator.Skip(t)

	for _, tc := range CommonColorCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := NewTestHelper(t)
			doc := h.CreateDocument(func(d *Document) {
				run := d.AddParagraph().AddRun()
				run.SetText("Colored")
				run.SetColor(tc.Input)
			})
			path := h.SaveDocument(doc, "color.docx")
			doc.Close()

			validator.AssertValid(t, path)
		})
	}
}

// =============================================================================
// Helper Functions
// =============================================================================

func formatFloat(f float64) string {
	if f == float64(int(f)) {
		return filepath.Base(string(rune(int(f) + '0')) + "pt")
	}
	return filepath.Base(string(rune(int(f)+30)) + "pt") // Unique identifier
}
