package presentation

import (
	"path/filepath"
	"strings"
	"testing"
)

type presentationFixtureCase struct {
	name   string
	mutate func(t *testing.T, pres Presentation)
	verify func(t *testing.T, pres Presentation)
}

func TestFixtureRoundTrip_Complex(t *testing.T) {
	cases := []presentationFixtureCase{
		{
			name: "minimal.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				tb := slide.AddTextBox(100000, 100000, 4000000, 800000)
				tb.SetText("Fixture Minimal")
				tb.SetPosition(120000, 140000)
				tb.SetSize(4100000, 900000)
				tf := tb.TextFrame()
				tf.SetAutofitType(AutofitNormal)
				paras := tf.Paragraphs()
				var para TextParagraph
				if len(paras) > 0 {
					para = paras[0]
				} else {
					para = tf.AddParagraph()
				}
				para.SetAlignment(AlignmentCenter)
				run := para.AddRun()
				run.SetText("Fixture Run")
				run.SetBold(true)
				run.SetItalic(true)
				run.SetUnderline(true)
				run.SetFontSize(18)
				run.SetFontName("Calibri")
				run.SetColor("FF0000")
				_ = slide.SetNotes("Fixture notes")
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if slide.Notes() != "Fixture notes" {
					t.Errorf("Notes() = %q, want Fixture notes", slide.Notes())
				}
				if !slideHasText(slide, "Fixture Minimal") {
					t.Error("Expected textbox after round-trip")
				}
				shape, err := slide.Shape("0")
				if err != nil {
					t.Fatalf("Shape(0) error = %v", err)
				}
				if shape.Left() != 120000 || shape.Top() != 140000 {
					t.Errorf("Position = (%d, %d), want (120000, 140000)", shape.Left(), shape.Top())
				}
				if shape.Width() != 4100000 || shape.Height() != 900000 {
					t.Errorf("Size = (%d, %d), want (4100000, 900000)", shape.Width(), shape.Height())
				}
				if shape.TextFrame().AutofitType() != AutofitNormal {
					t.Errorf("AutofitType() = %d, want %d", shape.TextFrame().AutofitType(), AutofitNormal)
				}
				var targetPara TextParagraph
				for _, para := range shape.TextFrame().Paragraphs() {
					if strings.Contains(para.Text(), "Fixture Run") {
						targetPara = para
						break
					}
				}
				if targetPara == nil {
					t.Fatal("Expected paragraph with Fixture Run after round-trip")
				}
				if targetPara.Alignment() != AlignmentCenter {
					t.Errorf("Alignment() = %d, want %d", targetPara.Alignment(), AlignmentCenter)
				}
				var targetRun TextRun
				for _, run := range targetPara.Runs() {
					if run.Text() == "Fixture Run" {
						targetRun = run
						break
					}
				}
				if targetRun == nil {
					t.Fatal("Expected run with Fixture Run after round-trip")
				}
				if !targetRun.Bold() || !targetRun.Italic() || !targetRun.Underline() {
					t.Error("Expected run formatting after round-trip")
				}
				if targetRun.FontSize() != 18 {
					t.Errorf("FontSize() = %v, want 18", targetRun.FontSize())
				}
				if targetRun.FontName() != "Calibri" {
					t.Errorf("FontName() = %q, want Calibri", targetRun.FontName())
				}
				if targetRun.Color() != "FF0000" {
					t.Errorf("Color() = %q, want FF0000", targetRun.Color())
				}
			},
		},
		{
			name: "title_slide.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if placeholder := slide.TitlePlaceholder(); placeholder != nil {
					_ = placeholder.SetText("Fixture Title")
				} else {
					tb := slide.AddTextBox(100000, 100000, 3000000, 800000)
					tb.SetText("Fixture Title")
				}
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if !slideHasText(slide, "Fixture Title") {
					t.Error("Expected title text after round-trip")
				}
				if placeholder := slide.TitlePlaceholder(); placeholder != nil && !placeholder.IsPlaceholder() {
					t.Error("Expected title placeholder after round-trip")
				}
			},
		},
		{
			name: "bullet_points.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				tb := slide.AddTextBox(100000, 100000, 3000000, 1200000)
				tf := tb.TextFrame()
				para := tf.AddParagraph()
				para.SetBulletType(BulletCharacter)
				para.SetText("Fixture Bullet")
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if !slideHasText(slide, "Fixture Bullet") {
					t.Error("Expected bullet text after round-trip")
				}
				if len(slide.Placeholders()) == 0 {
					t.Error("Expected placeholders after round-trip")
				}
			},
		},
		{
			name: "shapes.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				shape := slide.AddShape(ShapeTypeRectangle)
				shape.SetFillColor("00FF00")
				shape.SetLineColor("FF0000", 12700)
				shape.SetPosition(200000, 300000)
				shape.SetSize(1500000, 800000)
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if len(slide.Shapes()) == 0 {
					t.Error("Expected shapes after round-trip")
				}
				if shape, err := slide.Shape("0"); err == nil {
					if shape.Left() == 0 || shape.Top() == 0 {
						t.Error("Expected non-zero shape position after round-trip")
					}
					if shape.Width() == 0 || shape.Height() == 0 {
						t.Error("Expected non-zero shape size after round-trip")
					}
				}
			},
		},
		{
			name: "tables.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				table := slide.AddTable(2, 2, 100000, 100000, 3000000, 1000000)
				table.Cell(0, 0).SetText("Fixture Table")
				table.Row(0).SetHeight(500000)
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if len(slide.Tables()) == 0 {
					t.Error("Expected table after round-trip")
				}
				if len(slide.Tables()) > 0 && len(slide.Tables()[0].Rows()) > 0 {
					if slide.Tables()[0].Row(0).Height() == 0 {
						t.Error("Expected non-zero row height after round-trip")
					}
				}
			},
		},
		{
			name: "notes.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				_ = slide.SetNotes("Fixture Notes")
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if slide.Notes() != "Fixture Notes" {
					t.Errorf("Notes() = %q, want Fixture Notes", slide.Notes())
				}
			},
		},
		{
			name: "comments.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if _, err := slide.AddComment("Fixture Comment", "Tester", 100, 100); err != nil {
					t.Fatalf("AddComment() error = %v", err)
				}
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if len(slide.Comments()) == 0 {
					t.Error("Expected comments after round-trip")
				}
			},
		},
		{
			name: "hidden_slides.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide := pres.AddSlide(0)
				slide.SetHidden(true)
			},
			verify: func(t *testing.T, pres Presentation) {
				foundHidden := false
				for _, slide := range pres.Slides() {
					if slide.Hidden() {
						foundHidden = true
						break
					}
				}
				if !foundHidden {
					t.Error("Expected hidden slide after round-trip")
				}
			},
		},
		{
			name: "multiple_masters.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				pres.AddSlide(0)
			},
			verify: func(t *testing.T, pres Presentation) {
				if len(pres.Masters()) == 0 || len(pres.Layouts()) == 0 {
					t.Error("Expected masters/layouts after round-trip")
				}
			},
		},
		{
			name: "layouts.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				pres.AddSlide(0)
			},
			verify: func(t *testing.T, pres Presentation) {
				if len(pres.Layouts()) == 0 {
					t.Error("Expected layouts after round-trip")
				}
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if slide.Layout() == nil {
					t.Error("Expected layout after round-trip")
				}
			},
		},
		{
			name: "images.pptx",
			mutate: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if _, err := slide.AddPicture(filepath.Join("..", "..", "testdata", "pptx", "image1.png"), 200000, 200000, 1000000, 1000000); err != nil {
					t.Fatalf("AddPicture() error = %v", err)
				}
				if err := slide.ReplacePictureImage("0", filepath.Join("..", "..", "testdata", "pptx", "image1.png")); err != nil {
					t.Fatalf("ReplacePictureImage() error = %v", err)
				}
				slide.AddTextBox(100000, 100000, 3000000, 800000).SetText("Fixture Image Placeholder")
			},
			verify: func(t *testing.T, pres Presentation) {
				slide, err := pres.Slide(1)
				if err != nil {
					t.Fatalf("Slide(1) error = %v", err)
				}
				if !slideHasText(slide, "Fixture Image Placeholder") {
					t.Error("Expected placeholder text after round-trip")
				}
				if len(slide.Pictures()) == 0 {
					t.Error("Expected pictures after round-trip")
				}
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			pres, err := Open(fixtureRoundTripPath(tc.name))
			if err != nil {
				t.Fatalf("Open() error = %v", err)
			}
			if tc.mutate != nil {
				tc.mutate(t, pres)
			}

			outPath := filepath.Join(t.TempDir(), tc.name)
			if err := pres.SaveAs(outPath); err != nil {
				t.Fatalf("SaveAs() error = %v", err)
			}
			if err := pres.Close(); err != nil {
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

func slideHasText(slide Slide, text string) bool {
	for _, shape := range slide.Shapes() {
		if strings.Contains(shape.Text(), text) {
			return true
		}
	}
	return false
}

func fixtureRoundTripPath(name string) string {
	return filepath.Join("..", "..", "testdata", "pptx", name)
}
