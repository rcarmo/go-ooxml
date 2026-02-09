package e2e

import (
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestAdvancedCorporateFrankensteinArtifacts(t *testing.T) {
	artifacts := filepath.Join("..", "artifacts")
	if err := ensureArtifactsDir(artifacts); err != nil {
		t.Fatalf("ensureArtifactsDir() error = %v", err)
	}

	t.Run("word", func(t *testing.T) {
		out := filepath.Join(artifacts, "frankenstein_corporate_brief.docx")
		if err := buildCorporateFrankensteinDoc(out); err != nil {
			t.Fatalf("buildCorporateFrankensteinDoc() error = %v", err)
		}
	})
	t.Run("excel", func(t *testing.T) {
		out := filepath.Join(artifacts, "frankenstein_operating_model.xlsx")
		if err := buildCorporateFrankensteinWorkbook(out); err != nil {
			t.Fatalf("buildCorporateFrankensteinWorkbook() error = %v", err)
		}
	})
	t.Run("pptx", func(t *testing.T) {
		out := filepath.Join(artifacts, "frankenstein_exec_brief.pptx")
		if err := buildCorporateFrankensteinDeck(out); err != nil {
			t.Fatalf("buildCorporateFrankensteinDeck() error = %v", err)
		}
	})
}

func buildCorporateFrankensteinDoc(outPath string) error {
	doc, err := document.New()
	if err != nil {
		return err
	}
	defer doc.Close()

	doc.EnableTrackChanges("Mary Shelley")

	if err := setCorporateDocProperties(doc); err != nil {
		return err
	}

	doc.AddHeader(document.HeaderFooterDefault).SetText("Promethean Systems | Confidential")
	doc.AddFooter(document.HeaderFooterDefault).SetText("Frankenstein Initiative")

	titleStyle := doc.AddParagraphStyle("CorpTitle", "Corporate Title")
	titleStyle.SetBold(true)
	titleStyle.SetFontSize(28)
	titleStyle.SetFontName("Arial")
	titleStyle.SetColor("1F4E79")

	bodyStyle := doc.AddParagraphStyle("CorpBody", "Corporate Body")
	bodyStyle.SetFontSize(11)
	bodyStyle.SetFontName("Calibri")

	title := doc.AddParagraph()
	title.SetStyle("CorpTitle")
	title.AddRun().SetText("Frankenstein Initiative: Corporate Integration Brief")

	subtitle := doc.AddParagraph()
	subtitle.SetStyle("CorpBody")
	subtitle.AddRun().SetText("Prepared for the Promethean Systems Board")

	execSummary := doc.AddParagraph()
	execSummary.SetStyle("Heading1")
	execSummary.AddRun().SetText("Executive Summary")

	para := doc.AddParagraph()
	para.SetStyle("CorpBody")
	run := para.AddRun()
	run.SetText("In alignment with the Board’s directive, the Frankenstein Initiative consolidates disparate research assets into a single operational organism.")
	run.SetBold(true)
	run.SetColor("2F5597")

	risk := doc.AddParagraph()
	risk.SetStyle("CorpBody")
	risk.InsertTrackedText("Risk posture elevated. Governance gates required before full activation.")

	if _, err := doc.Comments().Add("Validate governance language with compliance", "Compliance", "Risk posture"); err != nil {
		return err
	}

	cc := doc.AddBlockContentControl("InitiativeName", "Initiative Name", "Frankenstein Initiative")
	if cc == nil {
		return err
	}

	scope := doc.AddParagraph()
	scope.SetStyle("Heading1")
	scope.AddRun().SetText("Scope and Operating Model")

	numID, err := doc.AddNumberedListStyle()
	if err != nil {
		return err
	}
	step1 := doc.AddParagraph()
	step1.SetStyle("CorpBody")
	_ = step1.SetList(numID, 0)
	step1.AddRun().SetText("Assemble cross-functional stewardship council.")

	step2 := doc.AddParagraph()
	step2.SetStyle("CorpBody")
	_ = step2.SetList(numID, 0)
	step2.AddRun().SetText("Deploy compliance laboratory for ethical review.")

	step3 := doc.AddParagraph()
	step3.SetStyle("CorpBody")
	_ = step3.SetList(numID, 0)
	step3.AddRun().SetText("Scale operations via regional implementation pods.")

	table := doc.AddTable(4, 4)
	table.SetStyle("TableGrid")
	table.Cell(0, 0).SetText("Division")
	table.Cell(0, 1).SetText("Owner")
	table.Cell(0, 2).SetText("Objective")
	table.Cell(0, 3).SetText("Status")
	table.Cell(1, 0).SetText("Laboratory")
	table.Cell(1, 1).SetText("Dr. Frankenstein")
	table.Cell(1, 2).SetText("Prototype stewardship")
	table.Cell(1, 3).SetText("In Progress")
	table.Cell(2, 0).SetText("Compliance")
	table.Cell(2, 1).SetText("Legal")
	table.Cell(2, 2).SetText("Ethical review")
	table.Cell(2, 3).SetText("Pending")
	table.Cell(3, 0).SetText("Operations")
	table.Cell(3, 1).SetText("Facilities")
	table.Cell(3, 2).SetText("Scaling readiness")
	table.Cell(3, 3).SetText("Planned")

	table.Cell(1, 0).SetGridSpan(2)
	table.Cell(2, 0).SetVerticalMerge(document.VerticalMerge("restart"))
	table.Cell(3, 0).SetVerticalMerge(document.VerticalMerge("continue"))

	deliverables := doc.AddParagraph()
	deliverables.SetStyle("Heading1")
	deliverables.AddRun().SetText("Deliverables and Timeline")

	fieldPara := doc.AddParagraph()
	fieldPara.SetStyle("CorpBody")
	if _, err := fieldPara.AddField("DATE", "2026-02-06"); err != nil {
		return err
	}

	linkPara := doc.AddParagraph()
	linkPara.SetStyle("CorpBody")
	if _, err := linkPara.AddHyperlink("https://example.com/frankenstein", "Program portal"); err != nil {
		return err
	}

	bookmarkPara := doc.AddParagraph()
	bookmarkPara.SetStyle("CorpBody")
	bookmarkPara.AddRun().SetText("Appendix: Corporate Covenant")
	if err := bookmarkPara.AddBookmark("CorpCovenant", 0, 0); err != nil {
		return err
	}
	jumpPara := doc.AddParagraph()
	jumpPara.SetStyle("CorpBody")
	if _, err := jumpPara.AddBookmarkLink("CorpCovenant", "Jump to covenant"); err != nil {
		return err
	}

	if _, err := doc.Body().AddChart(utils.InchesToEMU(6.0), utils.InchesToEMU(3.0), "KPI Trend Chart"); err != nil {
		return err
	}
	if _, err := doc.Body().AddDiagram(utils.InchesToEMU(6.0), utils.InchesToEMU(3.0), "Governance Diagram"); err != nil {
		return err
	}
	if _, err := doc.Body().AddPicture(filepath.Join("..", "testdata", "pptx", "image1.png"), utils.InchesToEMU(2.5), utils.InchesToEMU(2.0)); err != nil {
		return err
	}

	final := doc.AddParagraph()
	final.SetStyle("CorpBody")
	final.AddRun().SetText("Prepared with utmost discretion, for executive use only.")

	return doc.SaveAs(outPath)
}

func setCorporateDocProperties(doc document.Document) error {
	props := doc.Properties()
	props.Title = "Frankenstein Initiative Brief"
	props.Subject = "Corporate integration program"
	props.Creator = "Promethean Systems"
	props.Description = "Board-level brief for the Frankenstein Initiative"
	return doc.SetCoreProperties(&props)
}

func buildCorporateFrankensteinWorkbook(outPath string) error {
	wb, err := spreadsheet.New()
	if err != nil {
		return err
	}
	defer wb.Close()

	if err := setCorporateWorkbookProperties(wb); err != nil {
		return err
	}

	overview := wb.SheetsRaw()[0]
	_ = overview.SetName("Executive Summary")
	overview.Cell("A1").SetValue("Frankenstein Initiative")
	overview.Cell("A2").SetValue("Corporate KPIs")
	overview.Cell("A3").SetValue("Unit")
	overview.Cell("B3").SetValue("Target")
	overview.Cell("C3").SetValue("Actual")
	overview.Cell("A4").SetValue("Ethical Review Cycle")
	overview.Cell("B4").SetValue(30)
	overview.Cell("C4").SetValue(27)
	overview.Cell("A5").SetValue("Compliance Score")
	overview.Cell("B5").SetValue(95)
	overview.Cell("C5").SetFormula("C4+B5")
	overview.Cell("A6").SetValue("Operational Readiness")
	overview.Cell("B6").SetValue(80)
	overview.Cell("C6").SetValue(76)

	_ = overview.MergeCells("A1:C1")
	overview.Row(1).SetHeight(26)

	headerStyle := wb.Styles().Style().
		SetBold(true).
		SetFillColor("1F4E79")
	if err := overview.Cell("A3").SetStyle(headerStyle); err != nil {
		return err
	}
	if err := overview.Cell("B3").SetStyle(headerStyle); err != nil {
		return err
	}
	if err := overview.Cell("C3").SetStyle(headerStyle); err != nil {
		return err
	}

	_ = overview.Cell("C4").SetNumberFormat("0")
	_ = overview.Cell("C5").SetNumberFormat("0")
	_ = overview.Cell("C6").SetNumberFormat("0")

	if err := overview.Cell("A4").SetComment("Align with Geneva advisory cycle", "Compliance"); err != nil {
		return err
	}

	overviewTable := overview.AddTable("A3:C6", "ExecutiveKPIs")
	_ = overviewTable.AddRow(map[string]interface{}{
		"Column1": "Board Confidence",
		"Column2": 90,
		"Column3": 88,
	})
	if err := overview.AddChart("A8", "E20", "KPI Chart"); err != nil {
		return err
	}
	if err := overview.AddDiagram("A22", "C28", "Governance Diagram"); err != nil {
		return err
	}

	finance := wb.AddSheet("Budget")
	finance.Cell("A1").SetValue("Workstream")
	finance.Cell("B1").SetValue("Cost (USD)")
	finance.Cell("C1").SetValue("Status")
	finance.Cell("A2").SetValue("Assembly")
	finance.Cell("B2").SetValue(1250000)
	finance.Cell("C2").SetValue("Approved")
	finance.Cell("A3").SetValue("Compliance Lab")
	finance.Cell("B3").SetValue(750000)
	finance.Cell("C3").SetValue("Pending")
	finance.Cell("A4").SetValue("Deployment Pods")
	finance.Cell("B4").SetValue(2100000)
	finance.Cell("C4").SetValue("Planned")
	_ = finance.MergeCells("A1:C1")

	if err := finance.Cell("B2").SetNumberFormat("$#,##0"); err != nil {
		return err
	}
	if err := finance.Cell("B3").SetNumberFormat("$#,##0"); err != nil {
		return err
	}
	if err := finance.Cell("B4").SetNumberFormat("$#,##0"); err != nil {
		return err
	}

	finTable := finance.AddTable("A1:C4", "BudgetTable")
	_ = finTable.UpdateRow(1, map[string]interface{}{
		"Column1": "Assembly",
		"Column2": 1250000,
		"Column3": "Approved",
	})

	risk := wb.AddSheet("Risk Register")
	risk.Cell("A1").SetValue("Risk")
	risk.Cell("B1").SetValue("Likelihood")
	risk.Cell("C1").SetValue("Impact")
	risk.Cell("D1").SetValue("Mitigation")
	risk.Cell("A2").SetValue("Public backlash")
	risk.Cell("B2").SetValue("High")
	risk.Cell("C2").SetValue("Severe")
	risk.Cell("D2").SetValue("Proactive PR campaign")
	risk.Cell("A3").SetValue("Containment breach")
	risk.Cell("B3").SetValue("Medium")
	risk.Cell("C3").SetValue("High")
	risk.Cell("D3").SetValue("Facility hardening")
	_ = risk.Cell("A2").SetComment("Board visibility required", "CEO")

	riskTable := risk.AddTable("A1:D3", "RiskTable")
	_ = riskTable.AddRow(map[string]interface{}{
		"Column1": "IP exposure",
		"Column2": "Low",
		"Column3": "Medium",
		"Column4": "Patent strategy",
	})

	named := wb.AddNamedRange("RiskStatus", "'Risk Register'!$B$2:$B$4")
	named.SetHidden(true)

	hidden := wb.AddSheet("Archive")
	hidden.SetHidden(true)
	hidden.Cell("A1").SetValue("Do not distribute")

	return wb.SaveAs(outPath)
}

func setCorporateWorkbookProperties(wb spreadsheet.Workbook) error {
	props, err := wb.CoreProperties()
	if err != nil {
		return err
	}
	props.Title = "Frankenstein Initiative Operating Model"
	props.Subject = "Corporate KPI tracking"
	props.Creator = "Promethean Systems"
	props.Description = "Executive workbook for initiative oversight"
	return wb.SetCoreProperties(props)
}

func buildCorporateFrankensteinDeck(outPath string) error {
	pres, err := presentation.New()
	if err != nil {
		return err
	}
	defer pres.Close()

	if err := setCorporatePresentationProperties(pres); err != nil {
		return err
	}
	if err := pres.SetSlideSize(9144000, 6858000); err != nil {
		return err
	}

	titleSlide := pres.AddSlide(0)
	titleBox := titleSlide.AddTextBox(500000, 400000, 8000000, 1000000)
	titleBox.SetText("Frankenstein Initiative: Executive Alignment")
	titleBox.SetFillColor("1F4E79")

	subBox := titleSlide.AddTextBox(500000, 1500000, 8000000, 800000)
	subBox.SetText("Promethean Systems | Board Briefing")

	agenda := pres.AddSlide(0)
	agendaTitle := agenda.AddTextBox(500000, 300000, 8000000, 800000)
	agendaTitle.SetText("Agenda")
	body := agenda.AddTextBox(700000, 1300000, 7600000, 3500000)
	bodyTf := body.TextFrame()
	bodyTf.ClearParagraphs()

	p1 := bodyTf.AddParagraph()
	p1.SetText("1. Strategic Rationale")
	p1.SetBulletType(presentation.BulletAutoNumber)
	p1.SetLevel(0)

	p2 := bodyTf.AddParagraph()
	p2.SetText("2. Governance and Ethics")
	p2.SetBulletType(presentation.BulletAutoNumber)
	p2.SetLevel(0)

	p3 := bodyTf.AddParagraph()
	p3.SetText("3. Operating Model and KPIs")
	p3.SetBulletType(presentation.BulletAutoNumber)
	p3.SetLevel(0)

	metrics := pres.AddSlide(0)
	metricsTitle := metrics.AddTextBox(500000, 300000, 8000000, 800000)
	metricsTitle.SetText("Operational KPIs")
	table := metrics.AddTable(3, 3, 500000, 1500000, 8000000, 2000000)
	table.Cell(0, 0).SetText("Metric")
	table.Cell(0, 1).SetText("Target")
	table.Cell(0, 2).SetText("Actual")
	table.Cell(1, 0).SetText("Compliance Score")
	table.Cell(1, 1).SetText("95")
	table.Cell(1, 2).SetText("92")
	table.Cell(2, 0).SetText("Activation Readiness")
	table.Cell(2, 1).SetText("80")
	table.Cell(2, 2).SetText("76")

	notes := "Ethical review prioritized. Containment protocols updated per board directive."
	if err := metrics.SetNotes(notes); err != nil {
		return err
	}
	if _, err := metrics.AddComment("Confirm KPI values before circulation", "Strategy", 120, 160); err != nil {
		return err
	}
	if _, err := metrics.AddChart(700000, 3800000, 7600000, 2000000, "Initiative KPI Chart"); err != nil {
		return err
	}

	risk := pres.AddSlide(0)
	riskTitle := risk.AddTextBox(500000, 300000, 8000000, 800000)
	riskTitle.SetText("Risk Mitigation")
	riskBody := risk.AddTextBox(700000, 1300000, 7600000, 3500000)
	riskTf := riskBody.TextFrame()
	riskTf.ClearParagraphs()

	r1 := riskTf.AddParagraph()
	r1.SetText("Public reaction managed via corporate communications.")
	r1.SetBulletType(presentation.BulletCharacter)
	r1.SetLevel(0)
	r2 := riskTf.AddParagraph()
	r2.SetText("Laboratory access gated with biometric clearance.")
	r2.SetBulletType(presentation.BulletCharacter)
	r2.SetLevel(0)
	r3 := riskTf.AddParagraph()
	r3.SetText("Escalation protocols rehearsed quarterly.")
	r3.SetBulletType(presentation.BulletCharacter)
	r3.SetLevel(0)
	if _, err := risk.AddDiagram(500000, 950000, 2500000, 2200000, "Risk Diagram"); err != nil {
		return err
	}

	appendix := pres.AddSlide(0)
	appendix.SetHidden(true)
	appendix.AddTextBox(500000, 400000, 8000000, 800000).SetText("Appendix: Covenant Summary")
	appendix.AddTextBox(700000, 1400000, 7600000, 3000000).
		SetText("Frankenstein Initiative will operate within board-approved ethical and operational guardrails.")

	return pres.SaveAs(outPath)
}

func setCorporatePresentationProperties(pres presentation.Presentation) error {
	props, err := pres.CoreProperties()
	if err != nil {
		return err
	}
	props.Title = "Frankenstein Initiative Executive Brief"
	props.Subject = "Corporate presentation"
	props.Creator = "Promethean Systems"
	props.Description = "Board-level PowerPoint for the Frankenstein Initiative"
	return pres.SetCoreProperties(props)
}
