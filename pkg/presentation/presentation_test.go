package presentation

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
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

	p := testutil.NewResource(t, func() (Presentation, error) {
		return NewWithSize(customWidth, customHeight)
	})

	w, h := p.SlideSize()
	if w != customWidth || h != customHeight {
		t.Errorf("SlideSize() = (%d, %d), want (%d, %d)", w, h, customWidth, customHeight)
	}
}

func TestPresentation_CoreProperties(t *testing.T) {
	p := testutil.NewResource(t, New)

	props := &common.CoreProperties{
		Title:   "Presentation Title",
		Creator: "Presentation Author",
	}
	if err := p.SetCoreProperties(props); err != nil {
		t.Fatalf("SetCoreProperties() error = %v", err)
	}
	got, err := p.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties() error = %v", err)
	}
	if got.Title != props.Title {
		t.Errorf("Title = %q, want %q", got.Title, props.Title)
	}
	if got.Creator != props.Creator {
		t.Errorf("Creator = %q, want %q", got.Creator, props.Creator)
	}
}

func TestPresentation_MastersAndLayouts(t *testing.T) {
	p, err := Open("/workspace/testdata/pptx/comments.pptx")
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer p.Close()

	if len(p.Masters()) == 0 {
		t.Fatal("Masters() should return at least one master")
	}
	if len(p.Layouts()) == 0 {
		t.Fatal("Layouts() should return at least one layout")
	}
	if p.Masters()[0].Path() == "" {
		t.Error("Master Path() should not be empty")
	}
	if p.Layouts()[0].Path() == "" {
		t.Error("Layout Path() should not be empty")
	}
}

// =============================================================================
// Slide Management Tests
// =============================================================================

func TestAddSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	// Add first slide
	slide1 := p.AddSlide(0)
	if slide1 == nil {
		t.Fatal("AddSlide() returned nil")
	}
	if p.SlideCount() != 1 {
		t.Errorf("SlideCount() = %d, want 1", p.SlideCount())
	}
	if slide1.Index() != 1 {
		t.Errorf("slide1.Index() = %d, want 1", slide1.Index())
	}

	// Add second slide
	slide2 := p.AddSlide(0)
	if p.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", p.SlideCount())
	}
	if slide2.Index() != 2 {
		t.Errorf("slide2.Index() = %d, want 2", slide2.Index())
	}
}

func TestInsertSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	// Add two slides
	p.AddSlide(0)
	p.AddSlide(0)

	// Insert at beginning
	newSlide := p.InsertSlide(1, 0)
	if p.SlideCount() != 3 {
		t.Errorf("SlideCount() = %d, want 3", p.SlideCount())
	}
	if newSlide.Index() != 1 {
		t.Errorf("newSlide.Index() = %d, want 1", newSlide.Index())
	}

	// Verify all indices updated
	for i, slide := range p.SlidesRaw() {
		if slide.Index() != i+1 {
			t.Errorf("slide %d Index() = %d, want %d", i, slide.Index(), i+1)
		}
	}
}

func TestDeleteSlide(t *testing.T) {
	p := testutil.NewResource(t, New)

	p.AddSlide(0)
	p.AddSlide(0)
	p.AddSlide(0)

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

	slide1 := p.AddSlide(0)
	tb := slide1.AddTextBox(100, 100, 500, 200)
	tb.SetText("Original Text")

	duplicated := p.DuplicateSlide(1)

	if p.SlideCount() != 2 {
		t.Errorf("SlideCount() = %d, want 2", p.SlideCount())
	}

	if duplicated.Index() != 2 {
		t.Errorf("duplicated.Index() = %d, want 2", duplicated.Index())
	}
}

func TestReorderSlides(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide1 := p.AddSlide(0)
	slide2 := p.AddSlide(0)
	slide3 := p.AddSlide(0)

	slide1ID := slide1.ID()
	slide2ID := slide2.ID()
	slide3ID := slide3.ID()

	// Reverse order
	if err := p.ReorderSlides([]int{3, 2, 1}); err != nil {
		t.Fatalf("ReorderSlides() error = %v", err)
	}

	// Check new order
	slides := p.SlidesRaw()
	if slides[0].ID() != slide3ID {
		t.Errorf("slides[0].ID() = %s, want %s", slides[0].ID(), slide3ID)
	}
	if slides[1].ID() != slide2ID {
		t.Errorf("slides[1].ID() = %s, want %s", slides[1].ID(), slide2ID)
	}
	if slides[2].ID() != slide1ID {
		t.Errorf("slides[2].ID() = %s, want %s", slides[2].ID(), slide1ID)
	}

	// Invalid reorder
	if err := p.ReorderSlides([]int{1, 2}); err == nil {
		t.Error("ReorderSlides() with wrong length should error")
	}
}

// =============================================================================
// Slide Properties Tests
// =============================================================================

func TestSlideHidden(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide(0)

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

	slide := p.AddSlide(0)
	slide.AddTextBox(0, 0, 100, 100)
	slide.AddTextBox(0, 0, 100, 100)

	initialCount := len(slide.Shapes())
	if err := slide.DeleteShape("0"); err != nil {
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

	slide := p.AddSlide(0)

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
// Comments Tests
// =============================================================================

func TestSlideCommentsRoundTrip(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide(0)
	_, err := slide.AddComment("Needs review", "Test Author", 100, 200)
	if err != nil {
		t.Fatalf("AddComment() error = %v", err)
	}

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "comments-roundtrip.pptx")
	if err := p.SaveAs(path); err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	_ = p.Close()

	p2, err := Open(path)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer p2.Close()

	slide2, err := p2.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1) error = %v", err)
	}
	comments := slide2.Comments()
	if len(comments) != 1 {
		t.Fatalf("Comments() count = %d, want 1", len(comments))
	}
	if comments[0].Text() != "Needs review" {
		t.Errorf("Comment text = %q, want %q", comments[0].Text(), "Needs review")
	}
	if comments[0].Author() != "Test Author" {
		t.Errorf("Comment author = %q, want %q", comments[0].Author(), "Test Author")
	}
}

// =============================================================================
// Save/Load Tests
// =============================================================================

func TestSaveAs(t *testing.T) {
	p := testutil.NewResource(t, New)

	slide := p.AddSlide(0)
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

	slide1 := p.AddSlide(0)
	tb := slide1.AddTextBox(100000, 100000, 5000000, 1000000)
	tb.SetText("Slide 1 Text")

	slide2 := p.AddSlide(0)
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

	slide := p.AddSlide(0)
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

	slide := p.AddSlide(0)
	shape := slide.AddShape(ShapeTypeRectangle)

	shape.SetFillColor("0000FF")
	shape.SetLineColor("FF0000", 12700) // 1pt line
	// No assertion - just verify no panic

	shape.SetNoFill()
	// No assertion - just verify no panic
}
