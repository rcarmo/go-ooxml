package document

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

// ECMA-376 Validation Tests
// These tests validate our output against the official Open XML SDK validator.
// They verify compliance with ECMA-376 Part 1 (WordprocessingML) specifications.

// skipIfNoValidator skips the test if the .NET validator is not available.
func skipIfNoValidator(t *testing.T) string {
	t.Helper()
	
	// Check if dotnet is available
	dotnetRoot := os.Getenv("DOTNET_ROOT")
	if dotnetRoot == "" {
		dotnetRoot = "/home/linuxbrew/.linuxbrew/opt/dotnet/libexec"
	}
	
	validatorPath := "../../tools/validator/OoxmlValidator/bin/Release/net10.0/OoxmlValidator.dll"
	if _, err := os.Stat(validatorPath); os.IsNotExist(err) {
		t.Skip("OOXML validator not built; run 'make validate' first")
	}
	
	return validatorPath
}

// validateDocument runs the .NET validator on a document file.
func validateDocument(t *testing.T, validatorPath, docPath string) (bool, string) {
	t.Helper()
	
	dotnetRoot := os.Getenv("DOTNET_ROOT")
	if dotnetRoot == "" {
		dotnetRoot = "/home/linuxbrew/.linuxbrew/opt/dotnet/libexec"
	}
	
	cmd := exec.Command("dotnet", validatorPath, docPath, "--json")
	cmd.Env = append(os.Environ(), "DOTNET_ROOT="+dotnetRoot)
	
	output, err := cmd.CombinedOutput()
	if err != nil {
		// Exit code 2 means validation errors, not command failure
		if exitErr, ok := err.(*exec.ExitError); ok && exitErr.ExitCode() == 2 {
			return false, string(output)
		}
		t.Logf("Validator error: %v", err)
		return false, string(output)
	}
	
	return strings.Contains(string(output), `"valid": true`), string(output)
}

// =============================================================================
// ECMA-376 Compliance Tests
// =============================================================================

// TestECMA376_MinimalDocument verifies a minimal valid document structure.
// Per ECMA-376 §11.3.10, a WordprocessingML document requires:
// - [Content_Types].xml
// - _rels/.rels with relationship to main document
// - word/document.xml with w:document > w:body structure
func TestECMA376_MinimalDocument(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	docPath := filepath.Join(tmpDir, "minimal.docx")
	
	doc, err := New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}
	
	// Per spec, empty body is valid
	err = doc.SaveAs(docPath)
	doc.Close()
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	
	valid, output := validateDocument(t, validatorPath, docPath)
	if !valid {
		t.Errorf("Minimal document failed validation:\n%s", output)
	}
}

// TestECMA376_ParagraphWithText verifies basic paragraph structure.
// Per ECMA-376 §17.3.1.22, w:p contains w:pPr (optional) and content (w:r, etc.)
func TestECMA376_ParagraphWithText(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	docPath := filepath.Join(tmpDir, "paragraph.docx")
	
	doc, _ := New()
	para := doc.AddParagraph()
	para.SetText("Hello World")
	
	err := doc.SaveAs(docPath)
	doc.Close()
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	
	valid, output := validateDocument(t, validatorPath, docPath)
	if !valid {
		t.Errorf("Paragraph document failed validation:\n%s", output)
	}
}

// TestECMA376_ParagraphStyles verifies paragraph style references.
// Per ECMA-376 §17.3.1.27, w:pStyle references must use w:val attribute.
func TestECMA376_ParagraphStyles(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	
	styles := []string{"Heading1", "Heading2", "Heading3", "Normal", "Title"}
	
	for _, style := range styles {
		t.Run(style, func(t *testing.T) {
			docPath := filepath.Join(tmpDir, style+".docx")
			
			doc, _ := New()
			para := doc.AddParagraph()
			para.SetText("Test " + style)
			para.SetStyle(style)
			
			err := doc.SaveAs(docPath)
			doc.Close()
			if err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			
			valid, output := validateDocument(t, validatorPath, docPath)
			if !valid {
				t.Errorf("Style %q failed validation:\n%s", style, output)
			}
		})
	}
}

// TestECMA376_RunFormatting verifies run-level formatting.
// Per ECMA-376 §17.3.2, w:rPr contains formatting like w:b, w:i, w:u, w:sz.
func TestECMA376_RunFormatting(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	docPath := filepath.Join(tmpDir, "formatting.docx")
	
	doc, _ := New()
	para := doc.AddParagraph()
	
	// Bold
	r1 := para.AddRun()
	r1.SetText("Bold ")
	r1.SetBold(true)
	
	// Italic
	r2 := para.AddRun()
	r2.SetText("Italic ")
	r2.SetItalic(true)
	
	// Underline
	r3 := para.AddRun()
	r3.SetText("Underline ")
	r3.SetUnderline(true)
	
	// Font size (per ECMA-376 §17.3.2.38, w:sz uses half-points)
	r4 := para.AddRun()
	r4.SetText("Large ")
	r4.SetFontSize(24) // 24pt
	
	// Color
	r5 := para.AddRun()
	r5.SetText("Red")
	r5.SetColor("FF0000")
	
	err := doc.SaveAs(docPath)
	doc.Close()
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	
	valid, output := validateDocument(t, validatorPath, docPath)
	if !valid {
		t.Errorf("Formatting document failed validation:\n%s", output)
	}
}

// TestECMA376_FontSizes verifies font size values.
// Per ECMA-376 §17.3.2.38, w:sz/@w:val is in half-points (1pt = 2 half-points).
func TestECMA376_FontSizes(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	
	// Test various font sizes in points
	sizes := []float64{8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72}
	
	for _, size := range sizes {
		t.Run(strings.ReplaceAll(strings.TrimSuffix(strings.TrimSuffix(
			strings.Replace(string(rune(size)), ".", "_", 1), "0"), "_"), ".", "_"), func(t *testing.T) {
			docPath := filepath.Join(tmpDir, "size.docx")
			
			doc, _ := New()
			para := doc.AddParagraph()
			run := para.AddRun()
			run.SetText("Test")
			run.SetFontSize(size)
			
			err := doc.SaveAs(docPath)
			doc.Close()
			if err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			
			valid, output := validateDocument(t, validatorPath, docPath)
			if !valid {
				t.Errorf("Font size %.1fpt failed validation:\n%s", size, output)
			}
		})
	}
}

// TestECMA376_Table verifies table structure.
// Per ECMA-376 §17.4.38, w:tbl must contain w:tblPr, w:tblGrid, and w:tr elements.
func TestECMA376_Table(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	docPath := filepath.Join(tmpDir, "table.docx")
	
	doc, _ := New()
	
	// Create 3x3 table
	tbl := doc.AddTable(3, 3)
	
	// Fill with data
	for row := 0; row < 3; row++ {
		for col := 0; col < 3; col++ {
			tbl.Cell(row, col).SetText("Cell")
		}
	}
	
	err := doc.SaveAs(docPath)
	doc.Close()
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	
	valid, output := validateDocument(t, validatorPath, docPath)
	if !valid {
		t.Errorf("Table document failed validation:\n%s", output)
	}
}

// TestECMA376_TableDimensions verifies various table sizes.
func TestECMA376_TableDimensions(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	
	dimensions := []struct {
		rows, cols int
	}{
		{1, 1},
		{1, 5},
		{5, 1},
		{2, 2},
		{5, 5},
		{10, 3},
	}
	
	for _, dim := range dimensions {
		t.Run(string(rune('0'+dim.rows))+"x"+string(rune('0'+dim.cols)), func(t *testing.T) {
			docPath := filepath.Join(tmpDir, "table.docx")
			
			doc, _ := New()
			tbl := doc.AddTable(dim.rows, dim.cols)
			tbl.Cell(0, 0).SetText("Test")
			
			err := doc.SaveAs(docPath)
			doc.Close()
			if err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			
			valid, output := validateDocument(t, validatorPath, docPath)
			if !valid {
				t.Errorf("Table %dx%d failed validation:\n%s", dim.rows, dim.cols, output)
			}
		})
	}
}

// TestECMA376_MixedContent verifies documents with paragraphs and tables.
func TestECMA376_MixedContent(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	docPath := filepath.Join(tmpDir, "mixed.docx")
	
	doc, _ := New()
	
	// Title
	title := doc.AddParagraph()
	title.SetText("Document Title")
	title.SetStyle("Heading1")
	
	// Intro paragraph
	intro := doc.AddParagraph()
	intro.SetText("This is an introduction paragraph.")
	
	// Table
	tbl := doc.AddTable(2, 2)
	tbl.Cell(0, 0).SetText("A1")
	tbl.Cell(0, 1).SetText("B1")
	tbl.Cell(1, 0).SetText("A2")
	tbl.Cell(1, 1).SetText("B2")
	
	// Conclusion
	conclusion := doc.AddParagraph()
	conclusion.SetText("This is the conclusion.")
	
	err := doc.SaveAs(docPath)
	doc.Close()
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	
	valid, output := validateDocument(t, validatorPath, docPath)
	if !valid {
		t.Errorf("Mixed content document failed validation:\n%s", output)
	}
}

// TestECMA376_UnicodeContent verifies Unicode text handling.
func TestECMA376_UnicodeContent(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	docPath := filepath.Join(tmpDir, "unicode.docx")
	
	doc, _ := New()
	
	// Various Unicode content
	texts := []string{
		"English text",
		"日本語テキスト",       // Japanese
		"中文文本",            // Chinese
		"한국어 텍스트",        // Korean
		"Текст на русском",    // Russian
		"النص العربي",        // Arabic
		"🎉 Emoji test 🎊",   // Emoji
	}
	
	for _, text := range texts {
		para := doc.AddParagraph()
		para.SetText(text)
	}
	
	err := doc.SaveAs(docPath)
	doc.Close()
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	
	valid, output := validateDocument(t, validatorPath, docPath)
	if !valid {
		t.Errorf("Unicode document failed validation:\n%s", output)
	}
}

// TestECMA376_ParagraphAlignment verifies alignment values.
// Per ECMA-376 §17.3.1.13, w:jc/@w:val must be one of: start, end, center, both, etc.
func TestECMA376_ParagraphAlignment(t *testing.T) {
	validatorPath := skipIfNoValidator(t)
	tmpDir := t.TempDir()
	
	alignments := []string{"left", "center", "right", "both"}
	
	for _, align := range alignments {
		t.Run(align, func(t *testing.T) {
			docPath := filepath.Join(tmpDir, align+".docx")
			
			doc, _ := New()
			para := doc.AddParagraph()
			para.SetText("Aligned text")
			para.SetAlignment(align)
			
			err := doc.SaveAs(docPath)
			doc.Close()
			if err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			
			valid, output := validateDocument(t, validatorPath, docPath)
			if !valid {
				t.Errorf("Alignment %q failed validation:\n%s", align, output)
			}
		})
	}
}
