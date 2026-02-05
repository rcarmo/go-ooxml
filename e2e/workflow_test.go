package e2e

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func TestWordWorkflow_TechnicalReport(t *testing.T) {
	doc, err := document.New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}

	doc.EnableTrackChanges("Test Author")

	header := doc.AddHeader(document.HeaderFooterDefault)
	header.SetText("Confidential")
	footer := doc.AddFooter(document.HeaderFooterDefault)
	footer.SetText("Page 1")

	h1 := doc.AddParagraph()
	h1.SetStyle("Heading1")
	h1.AddRun().SetText("Technical Report")

	table := doc.AddTable(3, 2)
	table.Cell(0, 0).SetText("Customer")
	table.Cell(0, 1).SetText("[CUSTOMER_NAME]")
	table.Cell(1, 0).SetText("Project")
	table.Cell(1, 1).SetText("[PROJECT_NAME]")

	target := table.Cell(0, 1).Paragraphs()[0]
	target.InsertTrackedText("Acme Corp")

	if _, err := doc.Comments().Add("Verify customer name with legal", "Reviewer", "Acme Corp"); err != nil {
		t.Fatalf("Add comment error = %v", err)
	}

	cc := doc.AddBlockContentControl("Customer", "Customer Name", "Acme Corp")
	if cc == nil {
		t.Fatal("AddBlockContentControl returned nil")
	}

	path := filepath.Join(t.TempDir(), "technical_report.docx")
	if err := doc.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	doc2, err := document.Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
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
	if doc2.Header(document.HeaderFooterDefault) == nil || doc2.Header(document.HeaderFooterDefault).Text() != "Confidential" {
		t.Error("Header text missing after round-trip")
	}
	if doc2.Footer(document.HeaderFooterDefault) == nil || doc2.Footer(document.HeaderFooterDefault).Text() != "Page 1" {
		t.Error("Footer text missing after round-trip")
	}
	if doc2.ContentControlByTag("Customer") == nil {
		t.Error("Content control missing after round-trip")
	}
	tables := doc2.Tables()
	if len(tables) != 1 {
		t.Fatalf("Expected 1 table after round-trip, got %d", len(tables))
	}
	if !strings.Contains(tables[0].Cell(0, 1).Text(), "Acme Corp") {
		t.Error("Expected customer cell text after round-trip")
	}
}

func TestExcelWorkflow_DataWorkbook(t *testing.T) {
	wb, err := spreadsheet.New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}

	data := wb.SheetsRaw()[0]
	_ = data.SetName("Data")
	data.Cell("A1").SetValue("Item")
	data.Cell("B1").SetValue("Qty")
	data.Cell("C1").SetValue("Total")
	data.Cell("A2").SetValue("Widget")
	data.Cell("B2").SetValue(10)
	data.Cell("C2").SetFormula("B2*2")

	table := data.AddTable("A1:C2", "Sales")
	table.UpdateRow(1, map[string]interface{}{
		"Column1": "Widget",
		"Column2": 10,
		"Column3": 20,
	})

	_ = data.Cell("A2").SetComment("Check inventory", "Analyst")
	_ = data.MergeCells("A1:B1")

	wb.AddNamedRange("TotalQty", "'Data'!$B$2")

	style := wb.Styles().Style().SetBold(true).SetFillColor("FFDDDD")
	if err := data.Cell("A1").SetStyle(style); err != nil {
		t.Fatalf("SetStyle() error = %v", err)
	}

	summary := wb.AddSheet("Summary")
	summary.Cell("A1").SetFormula("SUM(Data!B2)")

	path := filepath.Join(t.TempDir(), "data_workbook.xlsx")
	if err := wb.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	if err := wb.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	wb2, err := spreadsheet.Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer wb2.Close()

	if wb2.SheetCount() != 2 {
		t.Errorf("SheetCount() = %d, want 2", wb2.SheetCount())
	}
	if len(wb2.NamedRanges()) != 1 {
		t.Errorf("NamedRanges() count = %d, want 1", len(wb2.NamedRanges()))
	}
	loadedTable, err := wb2.Table("Sales")
	if err != nil {
		t.Fatalf("Table(Sales) error = %v", err)
	}
	if len(loadedTable.Rows()) != 1 {
		t.Errorf("Table rows = %d, want 1", len(loadedTable.Rows()))
	}
	if loadedTable.Rows()[0].Cell("Column1").String() != "Widget" {
		t.Errorf("Table value = %q, want %q", loadedTable.Rows()[0].Cell("Column1").String(), "Widget")
	}
	data2, _ := wb2.SheetRaw("Data")
	if data2.Cell("C2").Formula() != "B2*2" {
		t.Errorf("Formula() = %q, want %q", data2.Cell("C2").Formula(), "B2*2")
	}
	if data2.MergedCells() == nil || len(data2.MergedCells()) == 0 {
		t.Error("Merged cells missing after round-trip")
	}
	if comment, ok := data2.Cell("A2").Comment(); !ok || comment.Text() != "Check inventory" {
		t.Error("Comment missing after round-trip")
	}
}

func TestPowerPointWorkflow_SalesDeck(t *testing.T) {
	pres, err := presentation.New()
	if err != nil {
		t.Fatalf("New() error = %v", err)
	}

	slide1 := pres.AddSlide(0)
	title := slide1.AddTextBox(100000, 100000, 5000000, 800000)
	title.SetText("Quarterly Update")

	table := slide1.AddTable(2, 2, 100000, 1200000, 4000000, 2000000)
	table.Cell(0, 0).SetText("Item")
	table.Cell(0, 1).SetText("Qty")
	table.Cell(1, 0).SetText("Widget")
	table.Cell(1, 1).SetText("10")

	if err := slide1.SetNotes("Review revenue numbers"); err != nil {
		t.Fatalf("SetNotes() error = %v", err)
	}
	if _, err := slide1.AddComment("Verify figures", "Reviewer", 100, 200); err != nil {
		t.Fatalf("AddComment() error = %v", err)
	}

	slide2 := pres.AddSlide(0)
	slide2.SetHidden(true)
	slide2.AddTextBox(100000, 100000, 3000000, 800000).SetText("Appendix")

	path := filepath.Join(t.TempDir(), "sales_deck.pptx")
	if err := pres.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	if err := pres.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	pres2, err := presentation.Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer pres2.Close()

	if pres2.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", pres2.SlideCount())
	}
	s1, _ := pres2.Slide(1)
	if s1.Notes() != "Review revenue numbers" {
		t.Errorf("Notes() = %q, want %q", s1.Notes(), "Review revenue numbers")
	}
	comments := s1.Comments()
	if len(comments) != 1 || comments[0].Text() != "Verify figures" {
		t.Error("Comments missing after round-trip")
	}
	tables := s1.Tables()
	if len(tables) != 1 || tables[0].Cell(1, 0).Text() != "Widget" {
		t.Error("Table content missing after round-trip")
	}
	s2, _ := pres2.Slide(2)
	if !s2.Hidden() {
		t.Error("Expected slide 2 to be hidden after round-trip")
	}
}

func TestWordWorkflow_TemplateUpdate(t *testing.T) {
	path := filepath.Join("..", "testdata", "word", "headers_footers.docx")
	doc, err := document.Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}

	doc.AddParagraph().SetText("Appendix A")
	table := doc.AddTable(1, 1)
	table.Cell(0, 0).SetText("Key")

	outPath := filepath.Join(t.TempDir(), "template_update.docx")
	if err := doc.SaveAs(outPath); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	doc2, err := document.Open(outPath)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer doc2.Close()

	foundParagraph := false
	for _, para := range doc2.Paragraphs() {
		if para.Text() == "Appendix A" {
			foundParagraph = true
			break
		}
	}
	if !foundParagraph {
		t.Error("Expected appended paragraph after round-trip")
	}

	foundTable := false
	for _, tbl := range doc2.Tables() {
		if tbl.RowCount() > 0 && tbl.ColumnCount() > 0 && tbl.Cell(0, 0).Text() == "Key" {
			foundTable = true
			break
		}
	}
	if !foundTable {
		t.Error("Expected appended table after round-trip")
	}
}

func TestExcelWorkflow_TemplateUpdate(t *testing.T) {
	path := filepath.Join("..", "testdata", "excel", "multiple_sheets.xlsx")
	wb, err := spreadsheet.Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}

	sheet := wb.AddSheet("Added")
	if err := sheet.Cell("A1").SetValue("Hello"); err != nil {
		t.Fatalf("SetValue() error = %v", err)
	}
	if err := sheet.Cell("B1").SetValue(5); err != nil {
		t.Fatalf("SetValue() error = %v", err)
	}
	if err := sheet.Cell("C1").SetFormula("B1*2"); err != nil {
		t.Fatalf("SetFormula() error = %v", err)
	}

	outPath := filepath.Join(t.TempDir(), "template_update.xlsx")
	if err := wb.SaveAs(outPath); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	if err := wb.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	wb2, err := spreadsheet.Open(outPath)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer wb2.Close()

	added, err := wb2.Sheet("Added")
	if err != nil {
		t.Fatalf("Sheet(Added) error = %v", err)
	}
	if added.Cell("A1").String() != "Hello" {
		t.Errorf("Added!A1 = %q, want %q", added.Cell("A1").String(), "Hello")
	}
	if added.Cell("C1").Formula() != "B1*2" {
		t.Errorf("Added!C1 formula = %q, want %q", added.Cell("C1").Formula(), "B1*2")
	}
}

func TestPowerPointWorkflow_TemplateUpdate(t *testing.T) {
	path := filepath.Join("..", "testdata", "pptx", "tables.pptx")
	pres, err := presentation.Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}

	originalCount := pres.SlideCount()
	slide := pres.AddSlide(0)
	slide.AddTextBox(100000, 100000, 3000000, 800000).SetText("Appendix")
	if err := slide.SetNotes("Added notes"); err != nil {
		t.Fatalf("SetNotes() error = %v", err)
	}

	outPath := filepath.Join(t.TempDir(), "template_update.pptx")
	if err := pres.SaveAs(outPath); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	if err := pres.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}

	pres2, err := presentation.Open(outPath)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer pres2.Close()

	if pres2.SlideCount() != originalCount+1 {
		t.Errorf("SlideCount() = %d, want %d", pres2.SlideCount(), originalCount+1)
	}
	last, err := pres2.Slide(pres2.SlideCount())
	if err != nil {
		t.Fatalf("Slide(last) error = %v", err)
	}
	if last.Notes() != "Added notes" {
		t.Errorf("Notes() = %q, want %q", last.Notes(), "Added notes")
	}
	foundText := false
	for _, shape := range last.Shapes() {
		if shape.Text() == "Appendix" {
			foundText = true
			break
		}
	}
	if !foundText {
		t.Error("Expected appended slide text after round-trip")
	}
}
