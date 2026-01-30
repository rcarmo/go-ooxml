package presentation

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// =============================================================================
// Creation Tests
// =============================================================================

func TestNew(t *testing.T) {
	p := testutil.NewResource(t, New)

	// Check default dimensions (4:3)
	w, h := p.SlideSize()
	if w != SlideWidth4x3 || h != SlideHeight4x3 {
		t.Errorf("SlideSize() = (%d, %d), want (%d, %d)", w, h, SlideWidth4x3, SlideHeight4x3)
	}

	// New presentation should have no slides
	if got := p.SlideCount(); got != 0 {
		t.Errorf("SlideCount() = %d, want 0", got)
	}
}

func TestNewWidescreen(t *testing.T) {
	p := testutil.NewResource(t, NewWidescreen)

	w, h := p.SlideSize()
	if w != SlideWidth16x9 || h != SlideHeight16x9 {
		t.Errorf("SlideSize() = (%d, %d), want (%d, %d)", w, h, SlideWidth16x9, SlideHeight16x9)
	}
}

func TestNewWithSize(t *testing.T) {
	customWidth := int64(7200000)  // 8 inches
	customHeight := int64(5400000) // 6 inches

	p := testutil.NewResource(t, func() (*Presentation, error) {
		return NewWithSize(customWidth, customHeight)
	})

	w, h := p.SlideSize()
	if w != customWidth || h != customHeight {
		t.Errorf("SlideSize() = (%d, %d), want (%d, %d)", w, h, customWidth, customHeight)
	}
}

// =============================================================================
// Slide Management Tests
// =============================================================================

func TestAddSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	// Add first slide
	slide1 := p.AddSlide()
	if slide1 == nil {
		t.Fatal("AddSlide() returned nil")
	}
	if p.SlideCount() != 1 {
		t.Errorf("SlideCount() = %d, want 1", p.SlideCount())
	}
	if slide1.Index() != 0 {
		t.Errorf("slide1.Index() = %d, want 0", slide1.Index())
	}

	// Add second slide
	slide2 := p.AddSlide()
	if p.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", p.SlideCount())
	}
	if slide2.Index() != 1 {
		t.Errorf("slide2.Index() = %d, want 1", slide2.Index())
	}
}

func TestInsertSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	// Add two slides
	p.AddSlide()
	p.AddSlide()

	// Insert at beginning
	newSlide := p.InsertSlide(0)
	if p.SlideCount() != 3 {
		t.Errorf("SlideCount() = %d, want 3", p.SlideCount())
	}
	if newSlide.Index() != 0 {
		t.Errorf("newSlide.Index() = %d, want 0", newSlide.Index())
	}

	// Verify all indices updated
	for i, slide := range p.Slides() {
		if slide.Index() != i {
			t.Errorf("slide %d Index() = %d, want %d", i, slide.Index(), i)
		}
	}
}

func TestDeleteSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	p.AddSlide()
	p.AddSlide()
	p.AddSlide()

	// Delete middle slide
	if err := p.DeleteSlide(1); err != nil {
		t.Fatalf("DeleteSlide(1) error = %v", err)
	}

	if p.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", p.SlideCount())
	}

	// Delete out of range should fail
	if err := p.DeleteSlide(10); err == nil {
		t.Error("DeleteSlide(10) should return error")
	}
}

func TestDuplicateSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide1 := p.AddSlide()
	tb := slide1.AddTextBox(100, 100, 500, 200)
	tb.SetText("Original Text")

	duplicated, err := p.DuplicateSlide(0)
	if err != nil {
		t.Fatalf("DuplicateSlide(0) error = %v", err)
	}

	if p.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", p.SlideCount())
	}

	if duplicated.Index() != 1 {
		t.Errorf("duplicated.Index() = %d, want 1", duplicated.Index())
	}
}

func TestReorderSlides(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide1 := p.AddSlide()
	slide2 := p.AddSlide()
	slide3 := p.AddSlide()

	slide1ID := slide1.ID()
	slide2ID := slide2.ID()
	slide3ID := slide3.ID()

	// Reverse order
	if err := p.ReorderSlides([]int{2, 1, 0}); err != nil {
		t.Fatalf("ReorderSlides() error = %v", err)
	}

	// Check new order
	slides := p.Slides()
	if slides[0].ID() != slide3ID {
		t.Errorf("slides[0].ID() = %d, want %d", slides[0].ID(), slide3ID)
	}
	if slides[1].ID() != slide2ID {
		t.Errorf("slides[1].ID() = %d, want %d", slides[1].ID(), slide2ID)
	}
	if slides[2].ID() != slide1ID {
		t.Errorf("slides[2].ID() = %d, want %d", slides[2].ID(), slide1ID)
	}

	// Invalid reorder
	if err := p.ReorderSlides([]int{0, 1}); err == nil {
		t.Error("ReorderSlides() with wrong length should error")
	}
}

// =============================================================================
// Slide Properties Tests
// =============================================================================

func TestSlideHidden(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide()

	// Default is visible
	if slide.Hidden() {
		t.Error("New slide should not be hidden")
	}

	slide.SetHidden(true)
	if !slide.Hidden() {
		t.Error("Slide should be hidden after SetHidden(true)")
	}

	slide.SetHidden(false)
	if slide.Hidden() {
		t.Error("Slide should not be hidden after SetHidden(false)")
	}
}

// =============================================================================
// Shape Tests
// =============================================================================

func TestDeleteShape(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide()
	slide.AddTextBox(0, 0, 100, 100)
	slide.AddTextBox(0, 0, 100, 100)

	initialCount := len(slide.Shapes())
	if err := slide.DeleteShape(0); err != nil {
		t.Fatalf("DeleteShape(0) error = %v", err)
	}

	if len(slide.Shapes()) != initialCount-1 {
		t.Errorf("Shape count = %d, want %d", len(slide.Shapes()), initialCount-1)
	}
}

// Shape/text tests live in parameterized_test.go

// =============================================================================
// Notes Tests
// =============================================================================

func TestSlideNotes(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide()

	// Initially no notes
	if slide.HasNotes() {
		t.Error("New slide should not have notes")
	}

	slide.SetNotes("Speaker notes here")
	if !slide.HasNotes() {
		t.Error("Slide should have notes after SetNotes")
	}

	if slide.Notes() != "Speaker notes here" {
		t.Errorf("Notes() = %q, want %q", slide.Notes(), "Speaker notes here")
	}

	slide.AppendNotes("Additional notes")
	if !strings.Contains(slide.Notes(), "Speaker notes here") || !strings.Contains(slide.Notes(), "Additional notes") {
		t.Errorf("AppendNotes failed: %q", slide.Notes())
	}
}

// =============================================================================
// Save/Load Tests
// =============================================================================

func TestSaveAs(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide()
	tb := slide.AddTextBox(100000, 100000, 5000000, 1000000)
	tb.SetText("Test Presentation")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "test.pptx")

	if err := p.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = p.Close()

	// Verify file exists
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Error("Output file was not created")
	}
}

func TestRoundTrip(t *testing.T) {
	// Create presentation
	p := testutil.NewResource(t, New)

	slide1 := p.AddSlide()
	tb := slide1.AddTextBox(100000, 100000, 5000000, 1000000)
	tb.SetText("Slide 1 Text")

	slide2 := p.AddSlide()
	slide2.SetHidden(true)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "roundtrip.pptx")

	if err := p.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = p.Close()

	// Reopen
	p2, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer p2.Close()

	// Verify structure
	if p2.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", p2.SlideCount())
	}
}

// =============================================================================
// Autofit Tests
// =============================================================================

func TestAutofit(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide()
	tb := slide.AddTextBox(0, 0, 1000000, 500000)
	tf := tb.TextFrame()

	// Default should be none
	if tf.AutofitType() != AutofitNone {
		t.Errorf("Default AutofitType() = %d, want AutofitNone", tf.AutofitType())
	}

	tf.SetAutofitType(AutofitNormal)
	if tf.AutofitType() != AutofitNormal {
		t.Errorf("AutofitType() = %d, want AutofitNormal", tf.AutofitType())
	}

	tf.SetAutofitType(AutofitShape)
	if tf.AutofitType() != AutofitShape {
		t.Errorf("AutofitType() = %d, want AutofitShape", tf.AutofitType())
	}
}

// =============================================================================
// Shape Fill Tests
// =============================================================================

func TestShapeFill(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide()
	shape := slide.AddShape(ShapeTypeRectangle, 0, 0, 1000000, 1000000)

	shape.SetFillColor("0000FF")
	shape.SetLineColor("FF0000", 12700) // 1pt line
	// No assertion - just verify no panic

	shape.SetNoFill()
	// No assertion - just verify no panic
}
