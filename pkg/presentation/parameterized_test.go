package presentation

import (
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// =============================================================================
// Test Fixtures
// =============================================================================

// PresentationFixture represents a test fixture for presentations.
type PresentationFixture struct {
	Name        string
	Description string
	Setup       func(*Presentation)
}

// CommonFixtures provides standard presentation test fixtures.
var CommonFixtures = []PresentationFixture{
	{
		Name:        "empty",
		Description: "Empty presentation with no slides",
		Setup:       func(p *Presentation) {},
	},
	{
		Name:        "single_slide",
		Description: "Single blank slide",
		Setup: func(p *Presentation) {
			p.AddSlide()
		},
	},
	{
		Name:        "multiple_slides",
		Description: "Three slides",
		Setup: func(p *Presentation) {
			p.AddSlide()
			p.AddSlide()
			p.AddSlide()
		},
	},
	{
		Name:        "slide_with_textbox",
		Description: "Slide with a text box",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			tb := s.AddTextBox(100000, 100000, 5000000, 1000000)
			tb.SetText("Test Text")
		},
	},
	{
		Name:        "slide_with_shapes",
		Description: "Slide with various shapes",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			s.AddShape(ShapeTypeRectangle, 100000, 100000, 1000000, 1000000)
			s.AddShape(ShapeTypeEllipse, 200000, 200000, 1000000, 1000000)
			s.AddTextBox(300000, 300000, 2000000, 500000)
		},
	},
	{
		Name:        "slide_with_notes",
		Description: "Slide with speaker notes",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			s.SetNotes("These are speaker notes")
		},
	},
	{
		Name:        "hidden_slide",
		Description: "Slide marked as hidden",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			s.SetHidden(true)
		},
	},
	{
		Name:        "formatted_text",
		Description: "Text with formatting",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			tb := s.AddTextBox(100000, 100000, 5000000, 2000000)
			tf := tb.TextFrame()
			
			p1 := tf.AddParagraph()
			r1 := p1.AddRun()
			r1.SetText("Bold")
			r1.SetBold(true)
			
			p2 := tf.AddParagraph()
			r2 := p2.AddRun()
			r2.SetText("Italic")
			r2.SetItalic(true)
			
			p3 := tf.AddParagraph()
			r3 := p3.AddRun()
			r3.SetText("Colored")
			r3.SetColor("FF0000")
		},
	},
	{
		Name:        "bullet_points",
		Description: "Slide with bullet points",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			tb := s.AddTextBox(100000, 100000, 5000000, 3000000)
			tf := tb.TextFrame()
			
			for i, text := range []string{"First point", "Second point", "Third point"} {
				para := tf.AddParagraph()
				para.SetBulletType(BulletCharacter)
				para.SetLevel(0)
				if i > 0 {
					para.SetLevel(1)
				}
				para.SetText(text)
			}
		},
	},
	{
		Name:        "slide_with_table",
		Description: "Slide with a table",
		Setup: func(p *Presentation) {
			s := p.AddSlide()
			table := s.AddTable(2, 2, 100000, 100000, 4000000, 2000000)
			table.Cell(0, 0).SetText("A1")
			table.Cell(0, 1).SetText("B1")
			table.Cell(1, 0).SetText("A2")
			table.Cell(1, 1).SetText("B2")
		},
	},
	{
		Name:        "widescreen",
		Description: "Widescreen 16:9 presentation",
		Setup:       func(p *Presentation) {}, // Created with NewWidescreen
	},
}

// =============================================================================
// Parameterized Tests
// =============================================================================

// TestFixtures_RoundTrip tests all fixtures with save/reload.
func TestFixtures_RoundTrip(t *testing.T) {
	for _, fixture := range CommonFixtures {
		t.Run(fixture.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			// Create and setup
			var p *Presentation
			var err error
			if fixture.Name == "widescreen" {
				p, err = NewWidescreen()
			} else {
				p, err = New()
			}
			h.RequireNoError(err, "New()")
			
			fixture.Setup(p)
			
			// Save
			path := h.TempFile(fixture.Name + ".pptx")
			h.RequireNoError(p.SaveAs(path), "SaveAs()")
			slideCount := p.SlideCount()
			p.Close()
			
			// Reload and verify
			p2, err := Open(path)
			h.RequireNoError(err, "Open()")
			defer p2.Close()
			
			h.AssertEqual(p2.SlideCount(), slideCount, "SlideCount after reload")
		})
	}
}

// =============================================================================
// Shape Type Parameterized Tests
// =============================================================================

// ShapeTypeTestCase represents a shape type test case.
type ShapeTypeTestCase struct {
	Name      string
	ShapeType ShapeType
}

var shapeTypeCases = []ShapeTypeTestCase{
	{"rectangle", ShapeTypeRectangle},
	{"ellipse", ShapeTypeEllipse},
	{"rounded_rect", ShapeTypeRoundRect},
	{"triangle", ShapeTypeTriangle},
	{"line", ShapeTypeLine},
	{"arrow", ShapeTypeArrow},
}

func TestShapeTypes_Parameterized(t *testing.T) {
	for _, tc := range shapeTypeCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			defer p.Close()
			
			slide := p.AddSlide()
			shape := slide.AddShape(tc.ShapeType, 0, 0, 1000000, 1000000)
			
			h.AssertNotNil(shape, "AddShape result")
			h.AssertEqual(shape.Type(), tc.ShapeType, "Shape type")
		})
	}
}

// =============================================================================
// Text Formatting Parameterized Tests
// =============================================================================

func TestTextFormatting_Parameterized(t *testing.T) {
	for _, fc := range testutil.CommonFormatCombinations {
		t.Run(fc.String(), func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			defer p.Close()
			
			slide := p.AddSlide()
			tb := slide.AddTextBox(0, 0, 2000000, 500000)
			tf := tb.TextFrame()
			para := tf.AddParagraph()
			run := para.AddRun()
			
			run.SetText("Test")
			if fc.Bold {
				run.SetBold(true)
			}
			if fc.Italic {
				run.SetItalic(true)
			}
			if fc.Underline {
				run.SetUnderline(true)
			}
			if fc.FontSize > 0 {
				run.SetFontSize(fc.FontSize)
			}
			if fc.FontName != "" {
				run.SetFontName(fc.FontName)
			}
			if fc.Color != "" {
				run.SetColor(fc.Color)
			}
			
			// Verify
			h.AssertEqual(run.Bold(), fc.Bold, "Bold")
			h.AssertEqual(run.Italic(), fc.Italic, "Italic")
			if fc.FontSize > 0 {
				h.AssertEqual(run.FontSize(), fc.FontSize, "FontSize")
			}
			if fc.FontName != "" {
				h.AssertEqual(run.FontName(), fc.FontName, "FontName")
			}
			if fc.Color != "" {
				h.AssertEqual(run.Color(), fc.Color, "Color")
			}
		})
	}
}

// =============================================================================
// Slide Operations Parameterized Tests
// =============================================================================

type slideOpTestCase struct {
	Name          string
	InitialSlides int
	Operation     func(*Presentation)
	WantSlides    int
}

var slideOpCases = []slideOpTestCase{
	{"add_to_empty", 0, func(p *Presentation) { p.AddSlide() }, 1},
	{"add_to_one", 1, func(p *Presentation) { p.AddSlide() }, 2},
	{"add_multiple", 0, func(p *Presentation) { p.AddSlide(); p.AddSlide(); p.AddSlide() }, 3},
	{"insert_at_start", 2, func(p *Presentation) { p.InsertSlide(0) }, 3},
	{"insert_at_end", 2, func(p *Presentation) { p.InsertSlide(2) }, 3},
	{"delete_first", 3, func(p *Presentation) { p.DeleteSlide(0) }, 2},
	{"delete_last", 3, func(p *Presentation) { p.DeleteSlide(2) }, 2},
	{"delete_middle", 3, func(p *Presentation) { p.DeleteSlide(1) }, 2},
	{"duplicate", 1, func(p *Presentation) { p.DuplicateSlide(0) }, 2},
}

func TestSlideOperations_Parameterized(t *testing.T) {
	for _, tc := range slideOpCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			defer p.Close()
			
			// Setup initial slides
			for i := 0; i < tc.InitialSlides; i++ {
				p.AddSlide()
			}
			
			// Perform operation
			tc.Operation(p)
			
			// Verify
			h.AssertEqual(p.SlideCount(), tc.WantSlides, "SlideCount")
		})
	}
}

// =============================================================================
// Alignment Parameterized Tests
// =============================================================================

type alignmentTestCase struct {
	Name      string
	Alignment Alignment
}

var alignmentCases = []alignmentTestCase{
	{"left", AlignmentLeft},
	{"center", AlignmentCenter},
	{"right", AlignmentRight},
	{"justify", AlignmentJustify},
}

func TestAlignment_Parameterized(t *testing.T) {
	for _, tc := range alignmentCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			defer p.Close()
			
			slide := p.AddSlide()
			tb := slide.AddTextBox(0, 0, 2000000, 500000)
			para := tb.TextFrame().AddParagraph()
			
			para.SetAlignment(tc.Alignment)
			h.AssertEqual(para.Alignment(), tc.Alignment, "Alignment")
		})
	}
}

// =============================================================================
// Bullet Type Parameterized Tests
// =============================================================================

type bulletTestCase struct {
	Name       string
	BulletType BulletType
}

var bulletCases = []bulletTestCase{
	{"none", BulletNone},
	{"auto_number", BulletAutoNumber},
	{"character", BulletCharacter},
}

func TestBulletTypes_Parameterized(t *testing.T) {
	for _, tc := range bulletCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			defer p.Close()
			
			slide := p.AddSlide()
			tb := slide.AddTextBox(0, 0, 2000000, 500000)
			para := tb.TextFrame().AddParagraph()
			
			para.SetBulletType(tc.BulletType)
			h.AssertEqual(para.BulletType(), tc.BulletType, "BulletType")
		})
	}
}

// =============================================================================
// Position/Size Parameterized Tests
// =============================================================================

type positionTestCase struct {
	Name   string
	Left   int64
	Top    int64
	Width  int64
	Height int64
}

var positionCases = []positionTestCase{
	{"origin", 0, 0, 1000000, 1000000},
	{"offset", 500000, 500000, 2000000, 1000000},
	{"large", 5000000, 3000000, 4000000, 2000000},
	{"small", 100000, 100000, 500000, 500000},
}

func TestShapePosition_Parameterized(t *testing.T) {
	for _, tc := range positionCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			defer p.Close()
			
			slide := p.AddSlide()
			shape := slide.AddTextBox(tc.Left, tc.Top, tc.Width, tc.Height)
			
			h.AssertEqual(shape.Left(), tc.Left, "Left")
			h.AssertEqual(shape.Top(), tc.Top, "Top")
			h.AssertEqual(shape.Width(), tc.Width, "Width")
			h.AssertEqual(shape.Height(), tc.Height, "Height")
		})
	}
}

// =============================================================================
// Round-Trip with String Data
// =============================================================================

func TestStringValues_RoundTrip(t *testing.T) {
	for _, tc := range testutil.CommonStringCases {
		t.Run(tc.Name, func(t *testing.T) {
			h := testutil.NewHelper(t)
			
			p, _ := New()
			slide := p.AddSlide()
			tb := slide.AddTextBox(0, 0, 2000000, 500000)
			tb.SetText(tc.Input)
			
			path := h.TempFile(tc.Name + ".pptx")
			h.RequireNoError(p.SaveAs(path), "SaveAs")
			p.Close()
			
			p2, _ := Open(path)
			defer p2.Close()
			
			shapes := p2.Slides()[0].Shapes()
			if len(shapes) > 0 {
				got := shapes[0].Text()
				h.AssertContains(got, tc.Want, "Text preserved")
			}
		})
	}
}
