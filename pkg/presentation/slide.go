package presentation

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
)

// Slide represents a slide in the presentation.
type Slide struct {
	pres  *Presentation
	slide *pml.Sld
	notes *pml.Notes
	id    int
	relID string
	index int
	path  string
}

// Index returns the 0-based index of the slide.
func (s *Slide) Index() int {
	return s.index
}

// ID returns the internal slide ID.
func (s *Slide) ID() int {
	return s.id
}

// Hidden returns whether the slide is hidden.
func (s *Slide) Hidden() bool {
	if s.slide.Show == nil {
		return false
	}
	return !*s.slide.Show
}

// SetHidden sets whether the slide is hidden.
func (s *Slide) SetHidden(hidden bool) {
	show := !hidden
	s.slide.Show = &show
}

// =============================================================================
// Shape access
// =============================================================================

// Shapes returns all shapes on the slide.
func (s *Slide) Shapes() []*Shape {
	if s.slide.CSld == nil || s.slide.CSld.SpTree == nil {
		return nil
	}

	var shapes []*Shape
	for _, item := range s.slide.CSld.SpTree.Content {
		if sp, ok := item.(*dml.Sp); ok {
			shapes = append(shapes, &Shape{
				slide: s,
				sp:    sp,
			})
		}
	}
	return shapes
}

// Shape returns a shape by name or index.
func (s *Slide) Shape(identifier interface{}) (*Shape, error) {
	shapes := s.Shapes()
	switch id := identifier.(type) {
	case int:
		if id < 0 || id >= len(shapes) {
			return nil, ErrShapeNotFound
		}
		return shapes[id], nil
	case string:
		for _, shape := range shapes {
			if shape.Name() == id {
				return shape, nil
			}
		}
		return nil, ErrShapeNotFound
	default:
		return nil, ErrShapeNotFound
	}
}

// AddTextBox adds a text box shape to the slide.
func (s *Slide) AddTextBox(left, top, width, height int64) *Shape {
	s.ensureSpTree()

	nextID := s.getNextShapeID()
	isTextBox := true

	sp := &dml.Sp{
		NvSpPr: &dml.NvSpPr{
			CNvPr: &dml.CNvPr{
				ID:   nextID,
				Name: "TextBox " + string(rune('0'+nextID)),
			},
			CNvSpPr: &dml.CNvSpPr{
				TxBox: &isTextBox,
			},
		},
		SpPr: &dml.SpPr{
			Xfrm: &dml.Xfrm{
				Off: &dml.Off{X: left, Y: top},
				Ext: &dml.Ext{Cx: width, Cy: height},
			},
			PrstGeom: &dml.PrstGeom{Prst: dml.PrstGeomRect},
		},
		TxBody: &dml.TxBody{
			BodyPr:   &dml.BodyPr{Wrap: "square"},
			LstStyle: &dml.LstStyle{},
			P:        []*dml.P{{EndParaRPr: &dml.RPr{}}},
		},
	}

	s.slide.CSld.SpTree.Content = append(s.slide.CSld.SpTree.Content, sp)

	return &Shape{slide: s, sp: sp}
}

// AddShape adds a shape of the specified type.
func (s *Slide) AddShape(shapeType ShapeType, left, top, width, height int64) *Shape {
	s.ensureSpTree()

	nextID := s.getNextShapeID()
	geomType := shapeTypeToGeom(shapeType)

	sp := &dml.Sp{
		NvSpPr: &dml.NvSpPr{
			CNvPr: &dml.CNvPr{
				ID:   nextID,
				Name: "Shape " + string(rune('0'+nextID)),
			},
			CNvSpPr: &dml.CNvSpPr{},
		},
		SpPr: &dml.SpPr{
			Xfrm: &dml.Xfrm{
				Off: &dml.Off{X: left, Y: top},
				Ext: &dml.Ext{Cx: width, Cy: height},
			},
			PrstGeom: &dml.PrstGeom{Prst: geomType},
		},
		TxBody: &dml.TxBody{
			BodyPr:   &dml.BodyPr{},
			LstStyle: &dml.LstStyle{},
			P:        []*dml.P{{EndParaRPr: &dml.RPr{}}},
		},
	}

	s.slide.CSld.SpTree.Content = append(s.slide.CSld.SpTree.Content, sp)

	return &Shape{slide: s, sp: sp}
}

// DeleteShape removes a shape from the slide.
func (s *Slide) DeleteShape(identifier interface{}) error {
	if s.slide.CSld == nil || s.slide.CSld.SpTree == nil {
		return ErrShapeNotFound
	}

	var indexToDelete int = -1

	switch id := identifier.(type) {
	case int:
		shapeIdx := 0
		for i, item := range s.slide.CSld.SpTree.Content {
			if _, ok := item.(*dml.Sp); ok {
				if shapeIdx == id {
					indexToDelete = i
					break
				}
				shapeIdx++
			}
		}
	case string:
		for i, item := range s.slide.CSld.SpTree.Content {
			if sp, ok := item.(*dml.Sp); ok {
				if sp.NvSpPr != nil && sp.NvSpPr.CNvPr != nil && sp.NvSpPr.CNvPr.Name == id {
					indexToDelete = i
					break
				}
			}
		}
	}

	if indexToDelete < 0 {
		return ErrShapeNotFound
	}

	s.slide.CSld.SpTree.Content = append(
		s.slide.CSld.SpTree.Content[:indexToDelete],
		s.slide.CSld.SpTree.Content[indexToDelete+1:]...,
	)

	return nil
}

// =============================================================================
// Placeholders
// =============================================================================

// Placeholders returns all placeholder shapes.
func (s *Slide) Placeholders() []*Shape {
	var placeholders []*Shape
	for _, shape := range s.Shapes() {
		if shape.IsPlaceholder() {
			placeholders = append(placeholders, shape)
		}
	}
	return placeholders
}

// TitlePlaceholder returns the title placeholder if present.
func (s *Slide) TitlePlaceholder() *Shape {
	for _, shape := range s.Shapes() {
		if shape.IsPlaceholder() {
			pt := shape.PlaceholderType()
			if pt == PlaceholderTitle || pt == PlaceholderCenteredTitle {
				return shape
			}
		}
	}
	return nil
}

// BodyPlaceholder returns the body placeholder if present.
func (s *Slide) BodyPlaceholder() *Shape {
	for _, shape := range s.Shapes() {
		if shape.IsPlaceholder() && shape.PlaceholderType() == PlaceholderBody {
			return shape
		}
	}
	return nil
}

// =============================================================================
// Notes
// =============================================================================

// Notes returns the notes text for this slide.
func (s *Slide) Notes() string {
	if s.notes == nil || s.notes.CSld == nil || s.notes.CSld.SpTree == nil {
		return ""
	}

	// Find text in notes shapes
	var text []string
	for _, item := range s.notes.CSld.SpTree.Content {
		if sp, ok := item.(*dml.Sp); ok {
			if sp.TxBody != nil {
				for _, p := range sp.TxBody.P {
					for _, r := range p.R {
						text = append(text, r.T)
					}
				}
			}
		}
	}
	return strings.Join(text, "\n")
}

// SetNotes sets the notes text for this slide.
func (s *Slide) SetNotes(text string) {
	s.ensureNotes()

	// Find or create a notes text shape
	var notesSp *dml.Sp
	for _, item := range s.notes.CSld.SpTree.Content {
		if sp, ok := item.(*dml.Sp); ok {
			if sp.NvSpPr != nil && sp.NvSpPr.NvPr != nil && sp.NvSpPr.NvPr.Ph != nil {
				if sp.NvSpPr.NvPr.Ph.Type == pml.PhTypeBody {
					notesSp = sp
					break
				}
			}
		}
	}

	if notesSp == nil {
		// Create a notes body placeholder
		notesSp = &dml.Sp{
			NvSpPr: &dml.NvSpPr{
				CNvPr:   &dml.CNvPr{ID: 2, Name: "Notes Placeholder"},
				CNvSpPr: &dml.CNvSpPr{},
				NvPr:    &dml.NvPr{Ph: &dml.Ph{Type: pml.PhTypeBody}},
			},
			SpPr: &dml.SpPr{},
			TxBody: &dml.TxBody{
				BodyPr:   &dml.BodyPr{},
				LstStyle: &dml.LstStyle{},
			},
		}
		s.notes.CSld.SpTree.Content = append(s.notes.CSld.SpTree.Content, notesSp)
	}

	// Set the text
	if notesSp.TxBody == nil {
		notesSp.TxBody = &dml.TxBody{
			BodyPr:   &dml.BodyPr{},
			LstStyle: &dml.LstStyle{},
		}
	}

	notesSp.TxBody.P = []*dml.P{{
		R: []*dml.R{{T: text}},
	}}
}

// AppendNotes appends text to the existing notes.
func (s *Slide) AppendNotes(text string) {
	existing := s.Notes()
	if existing != "" {
		text = existing + "\n" + text
	}
	s.SetNotes(text)
}

// HasNotes returns true if the slide has notes.
func (s *Slide) HasNotes() bool {
	return s.Notes() != ""
}

// =============================================================================
// Convenience methods
// =============================================================================

// Title returns the text of the title placeholder.
func (s *Slide) Title() string {
	if placeholder := s.TitlePlaceholder(); placeholder != nil {
		return placeholder.Text()
	}
	return ""
}

// SetTitle sets the text of the title placeholder.
func (s *Slide) SetTitle(text string) error {
	if placeholder := s.TitlePlaceholder(); placeholder != nil {
		return placeholder.SetText(text)
	}
	return ErrShapeNotFound
}

// =============================================================================
// Internal methods
// =============================================================================

func (s *Slide) ensureSpTree() {
	if s.slide.CSld == nil {
		s.slide.CSld = &pml.CSld{}
	}
	if s.slide.CSld.SpTree == nil {
		s.slide.CSld.SpTree = &pml.SpTree{
			NvGrpSpPr: &pml.NvGrpSpPr{
				CNvPr:      &pml.CNvPr{ID: 1, Name: ""},
				CNvGrpSpPr: &pml.CNvGrpSpPr{},
				NvPr:       &pml.NvPr{},
			},
			GrpSpPr: &pml.GrpSpPr{},
		}
	}
}

func (s *Slide) ensureNotes() {
	if s.notes == nil {
		s.notes = &pml.Notes{
			CSld: &pml.CSld{
				SpTree: &pml.SpTree{
					NvGrpSpPr: &pml.NvGrpSpPr{
						CNvPr:      &pml.CNvPr{ID: 1, Name: ""},
						CNvGrpSpPr: &pml.CNvGrpSpPr{},
						NvPr:       &pml.NvPr{},
					},
					GrpSpPr: &pml.GrpSpPr{},
				},
			},
		}
	}
}

func (s *Slide) getNextShapeID() int {
	maxID := 1
	if s.slide.CSld != nil && s.slide.CSld.SpTree != nil {
		for _, item := range s.slide.CSld.SpTree.Content {
			if sp, ok := item.(*dml.Sp); ok {
				if sp.NvSpPr != nil && sp.NvSpPr.CNvPr != nil && sp.NvSpPr.CNvPr.ID > maxID {
					maxID = sp.NvSpPr.CNvPr.ID
				}
			}
		}
	}
	return maxID + 1
}
