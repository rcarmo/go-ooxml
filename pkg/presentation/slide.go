package presentation

import (
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
)

// Slide represents a slide in the presentation.
type slideImpl struct {
	pres  *presentationImpl
	slide *pml.Sld
	notes *pml.Notes
	id    int
	relID string
	index int
	path  string
}

// Index returns the 1-based index of the slide.
func (s *slideImpl) Index() int {
	return s.index + 1
}

// ID returns the internal slide ID.
func (s *slideImpl) ID() string {
	return strconv.Itoa(s.id)
}

// Hidden returns whether the slide is hidden.
func (s *slideImpl) Hidden() bool {
	if s.slide.Show == nil {
		return false
	}
	return !*s.slide.Show
}

// SetHidden sets whether the slide is hidden.
func (s *slideImpl) SetHidden(hidden bool) {
	show := !hidden
	s.slide.Show = &show
}

// =============================================================================
// Shape access
// =============================================================================

// Shapes returns all shapes on the slide.
func (s *slideImpl) Shapes() []Shape {
	if s.slide.CSld == nil || s.slide.CSld.SpTree == nil {
		return nil
	}

	var shapes []Shape
	for _, item := range s.slide.CSld.SpTree.Content {
		if sp, ok := item.(*dml.Sp); ok {
			shapes = append(shapes, &shapeImpl{
				slide: s,
				sp:    sp,
			})
		}
		if gf, ok := item.(*pml.GraphicFrame); ok {
			shapes = append(shapes, &shapeImpl{
				slide: s,
				graphicFrame: gf,
			})
		}
	}
	return shapes
}

// Shape returns a shape by name or index.
func (s *slideImpl) Shape(identifier string) (Shape, error) {
	shapes := s.Shapes()
	if idx, err := strconv.Atoi(identifier); err == nil {
		if idx < 0 || idx >= len(shapes) {
			return nil, ErrShapeNotFound
		}
		return shapes[idx], nil
	}
	for _, shape := range shapes {
		if shape.Name() == identifier {
			return shape, nil
		}
	}
	return nil, ErrShapeNotFound
}

// Tables returns all tables on the slide.
func (s *slideImpl) Tables() []Table {
	var tables []Table
	for _, shape := range s.Shapes() {
		if shape.HasTable() {
			tables = append(tables, shape.Table())
		}
	}
	return tables
}

// AddTextBox adds a text box shape to the slide.
func (s *slideImpl) AddTextBox(left, top, width, height int64) Shape {
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

	return &shapeImpl{slide: s, sp: sp}
}

// AddShape adds a shape of the specified type.
func (s *slideImpl) addShape(shapeType ShapeType, left, top, width, height int64) Shape {
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

	return &shapeImpl{slide: s, sp: sp}
}

// AddShape adds a shape with default dimensions.
func (s *slideImpl) AddShape(shapeType ShapeType) Shape {
	return s.addShape(shapeType, 0, 0, 1000000, 1000000)
}

// AddPicture adds a picture placeholder (not implemented).
func (s *slideImpl) AddPicture(imagePath string, left, top, width, height int64) (Shape, error) {
	return nil, ErrInvalidIndex
}

// Comments returns slide comments (not implemented).
func (s *slideImpl) Comments() []Comment {
	return nil
}

// AddComment adds a slide comment (not implemented).
func (s *slideImpl) AddComment(text, author string, x, y float64) (Comment, error) {
	return nil, ErrInvalidIndex
}

// AddTable adds a table shape to the slide.
func (s *slideImpl) AddTable(rows, cols int, left, top, width, height int64) Table {
	s.ensureSpTree()

	nextID := s.getNextShapeID()
	table := newTable(rows, cols, width, height)

	gf := &pml.GraphicFrame{
		NvGraphicFramePr: &pml.NvGraphicFramePr{
			CNvPr: &pml.CNvPr{
				ID:   nextID,
				Name: "Table " + string(rune('0'+nextID)),
			},
			CNvGraphicFramePr: &pml.CNvGraphicFramePr{},
			NvPr: &pml.NvPr{Ph: &pml.Ph{Type: pml.PhTypeTbl}},
		},
		Xfrm: &pml.Xfrm{
			Off: &pml.Off{X: left, Y: top},
			Ext: &pml.Ext{Cx: width, Cy: height},
		},
		Graphic: &dml.Graphic{
			GraphicData: &dml.GraphicData{
				URI: dml.GraphicDataURITable,
				Tbl: table.tbl,
			},
		},
	}

	s.slide.CSld.SpTree.Content = append(s.slide.CSld.SpTree.Content, gf)

	return table
}


// DeleteShape removes a shape from the slide.
func (s *slideImpl) DeleteShape(identifier string) error {
	if s.slide.CSld == nil || s.slide.CSld.SpTree == nil {
		return ErrShapeNotFound
	}

	var indexToDelete int = -1

	if idx, err := strconv.Atoi(identifier); err == nil {
		shapeIdx := 0
		for i, item := range s.slide.CSld.SpTree.Content {
			if _, ok := item.(*dml.Sp); ok {
				if shapeIdx == idx {
					indexToDelete = i
					break
				}
				shapeIdx++
			}
			if _, ok := item.(*pml.GraphicFrame); ok {
				if shapeIdx == idx {
					indexToDelete = i
					break
				}
				shapeIdx++
			}
		}
	} else {
		for i, item := range s.slide.CSld.SpTree.Content {
			if sp, ok := item.(*dml.Sp); ok {
				if sp.NvSpPr != nil && sp.NvSpPr.CNvPr != nil && sp.NvSpPr.CNvPr.Name == identifier {
					indexToDelete = i
					break
				}
			}
			if gf, ok := item.(*pml.GraphicFrame); ok {
				if gf.NvGraphicFramePr != nil && gf.NvGraphicFramePr.CNvPr != nil && gf.NvGraphicFramePr.CNvPr.Name == identifier {
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
func (s *slideImpl) Placeholders() []Shape {
	var placeholders []Shape
	for _, shape := range s.Shapes() {
		if shape.IsPlaceholder() {
			placeholders = append(placeholders, shape)
		}
	}
	return placeholders
}

// TitlePlaceholder returns the title placeholder if present.
func (s *slideImpl) TitlePlaceholder() Shape {
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
func (s *slideImpl) BodyPlaceholder() Shape {
	for _, shape := range s.Shapes() {
		if shape.IsPlaceholder() && shape.PlaceholderType() == PlaceholderBody {
			return shape
		}
	}
	return nil
}

// Layout returns the slide layout (not implemented).
func (s *slideImpl) Layout() SlideLayout {
	return nil
}

// =============================================================================
// Notes
// =============================================================================

// Notes returns the notes text for this slide.
func (s *slideImpl) Notes() string {
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
func (s *slideImpl) SetNotes(text string) error {
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
	return nil
}

// AppendNotes appends text to the existing notes.
func (s *slideImpl) AppendNotes(text string) error {
	existing := s.Notes()
	if existing != "" {
		text = existing + "\n" + text
	}
	return s.SetNotes(text)
}

// HasNotes returns true if the slide has notes.
func (s *slideImpl) HasNotes() bool {
	return s.Notes() != ""
}

// =============================================================================
// Convenience methods
// =============================================================================

// Title returns the text of the title placeholder.
func (s *slideImpl) Title() string {
	if placeholder := s.TitlePlaceholder(); placeholder != nil {
		if sp, ok := placeholder.(*shapeImpl); ok {
			return sp.Text()
		}
	}
	return ""
}

// SetTitle sets the text of the title placeholder.
func (s *slideImpl) SetTitle(text string) error {
	if placeholder := s.TitlePlaceholder(); placeholder != nil {
		if sp, ok := placeholder.(*shapeImpl); ok {
			return sp.SetText(text)
		}
	}
	return ErrShapeNotFound
}

// =============================================================================
// Internal methods
// =============================================================================

func (s *slideImpl) ensureSpTree() {
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

func (s *slideImpl) ensureNotes() {
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

func (s *slideImpl) getNextShapeID() int {
	maxID := 1
	if s.slide.CSld != nil && s.slide.CSld.SpTree != nil {
		for _, item := range s.slide.CSld.SpTree.Content {
			if sp, ok := item.(*dml.Sp); ok {
				if sp.NvSpPr != nil && sp.NvSpPr.CNvPr != nil && sp.NvSpPr.CNvPr.ID > maxID {
					maxID = sp.NvSpPr.CNvPr.ID
				}
			}
			if gf, ok := item.(*pml.GraphicFrame); ok {
				if gf.NvGraphicFramePr != nil && gf.NvGraphicFramePr.CNvPr != nil && gf.NvGraphicFramePr.CNvPr.ID > maxID {
					maxID = gf.NvGraphicFramePr.CNvPr.ID
				}
			}
		}
	}
	return maxID + 1
}
