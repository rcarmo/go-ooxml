package presentation

import (
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/chart"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/diagram"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Slide represents a slide in the presentation.
type slideImpl struct {
	pres  *presentationImpl
	slide *pml.Sld
	notes *pml.Notes
	comments *pml.CommentList
	id    int
	relID string
	index int
	path  string
}

type commentImpl struct {
	slide   *slideImpl
	comment *pml.Comment
}

// ID returns the comment ID.
func (c *commentImpl) ID() string {
	if c == nil || c.comment == nil {
		return ""
	}
	return c.comment.ID
}

// Author returns the comment author.
func (c *commentImpl) Author() string {
	if c == nil || c.slide == nil || c.slide.pres == nil || c.comment == nil {
		return ""
	}
	return c.slide.pres.commentAuthorName(c.comment.AuthorID)
}

// Text returns the comment text.
func (c *commentImpl) Text() string {
	if c == nil || c.comment == nil || c.comment.TxBody == nil {
		return ""
	}
	var parts []string
	for _, p := range c.comment.TxBody.P {
		for _, r := range p.R {
			parts = append(parts, r.T)
		}
	}
	return strings.Join(parts, "")
}

// SetText sets the comment text.
func (c *commentImpl) SetText(text string) {
	if c == nil || c.comment == nil {
		return
	}
	if c.comment.TxBody == nil {
		c.comment.TxBody = &pml.CommentText{
			BodyPr:   &pml.CommentBodyPr{},
			LstStyle: &pml.CommentLstStyle{},
		}
	}
	c.comment.TxBody.P = []*pml.CommentParagraph{{
		R: []*pml.CommentRun{{
			RPr: &pml.CommentRunProps{Lang: "en-US"},
			T:   text,
		}},
	}}
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
		switch v := item.(type) {
		case *dml.Sp:
			shapes = append(shapes, &shapeImpl{
				slide: s,
				sp:    v,
			})
		case *pml.GraphicFrame:
			shapes = append(shapes, &shapeImpl{
				slide: s,
				graphicFrame: v,
			})
		case *pml.Pic:
			shapes = append(shapes, &shapeImpl{
				slide: s,
				pic:   v,
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
			NvPr: &dml.NvPr{},
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
	if s == nil || s.pres == nil {
		return nil, ErrInvalidIndex
	}
	s.ensureSpTree()

	target, err := s.pres.addImagePart(imagePath)
	if err != nil {
		return nil, err
	}
	sourcePath := s.path
	if sourcePath == "" {
		sourcePath = fmt.Sprintf("ppt/slides/slide%d.xml", s.index+1)
		s.path = sourcePath
	}
	rels := s.pres.pkg.GetRelationships(sourcePath)
	relID := rels.NextID()
	rels.AddWithID(relID, packaging.RelTypeImage, target, packaging.TargetModeInternal)
	nextID := s.getNextShapeID()
	noChange := true
	rotWithShape := true
	picName := fmt.Sprintf("Picture %d", nextID)
	pic := &pml.Pic{
		NvPicPr: &pml.NvPicPr{
			CNvPr: &pml.CNvPr{
				ID:   nextID,
				Name: picName,
			},
			CNvPicPr: &pml.CNvPicPr{
				PicLocks: &pml.PicLocks{
					NoChangeAspect: &noChange,
				},
			},
			NvPr: &pml.NvPr{},
		},
		BlipFill: &dml.BlipFill{
			RotWithShape: &rotWithShape,
			Blip: &dml.Blip{
				Embed: relID,
			},
			Stretch: &dml.Stretch{
				FillRect: &dml.FillRect{},
			},
		},
		SpPr: &dml.SpPr{
			Xfrm: &dml.Xfrm{
				Off: &dml.Off{X: left, Y: top},
				Ext: &dml.Ext{Cx: width, Cy: height},
			},
			PrstGeom: &dml.PrstGeom{Prst: dml.PrstGeomRect},
		},
	}

	s.slide.CSld.SpTree.Content = append(s.slide.CSld.SpTree.Content, pic)
	return &shapeImpl{slide: s, pic: pic}, nil
}

// AddChart adds a chart to the slide.
func (s *slideImpl) AddChart(left, top, width, height int64, title string) (Shape, error) {
	if s == nil || s.pres == nil {
		return nil, ErrInvalidIndex
	}
	s.ensureSpTree()

	nextID := s.getNextShapeID()
	cs := &chart.ChartSpace{
		Chart: &chart.Chart{
			PlotArea: &chart.PlotArea{Layout: &chart.Layout{}},
			Title:    &chart.Title{},
			Legend:   &chart.Legend{},
		},
	}
	data, err := utils.MarshalXMLWithHeader(cs)
	if err != nil {
		return nil, err
	}
	chartPath := fmt.Sprintf("ppt/charts/chart%d.xml", nextID)
	if _, err := s.pres.pkg.AddPart(chartPath, packaging.ContentTypeChart, data); err != nil {
		return nil, err
	}

	sourcePath := s.path
	if sourcePath == "" {
		sourcePath = fmt.Sprintf("ppt/slides/slide%d.xml", s.index+1)
		s.path = sourcePath
	}
	rels := s.pres.pkg.GetRelationships(sourcePath)
	relID := rels.NextID()
	rels.AddWithID(relID, packaging.RelTypeChart, relativeTarget(sourcePath, chartPath), packaging.TargetModeInternal)

	name := title
	if name == "" {
		name = fmt.Sprintf("Chart %d", nextID)
	}
	gf := &pml.GraphicFrame{
		NvGraphicFramePr: &pml.NvGraphicFramePr{
			CNvPr: &pml.CNvPr{
				ID:   nextID,
				Name: name,
			},
			CNvGraphicFramePr: &pml.CNvGraphicFramePr{},
			NvPr:              &pml.NvPr{Ph: &pml.Ph{Type: pml.PhTypeChart}},
		},
		Xfrm: &pml.Xfrm{
			Off: &pml.Off{X: left, Y: top},
			Ext: &pml.Ext{Cx: width, Cy: height},
		},
		Graphic: &dml.Graphic{
			GraphicData: &dml.GraphicData{
				URI:   dml.GraphicDataURIChart,
				Chart: &dml.ChartRef{RID: relID},
			},
		},
	}
	s.slide.CSld.SpTree.Content = append(s.slide.CSld.SpTree.Content, gf)
	return &shapeImpl{slide: s, graphicFrame: gf}, nil
}

// AddDiagram adds a diagram (SmartArt) to the slide.
func (s *slideImpl) AddDiagram(left, top, width, height int64, title string) (Shape, error) {
	if s == nil || s.pres == nil {
		return nil, ErrInvalidIndex
	}
	s.ensureSpTree()

	dataModel := &diagram.DataModel{}
	data, err := utils.MarshalXMLWithHeader(dataModel)
	if err != nil {
		return nil, err
	}
	layoutData, err := utils.MarshalXMLWithHeader(&diagram.LayoutDef{})
	if err != nil {
		return nil, err
	}
	styleData, err := utils.MarshalXMLWithHeader(&diagram.StyleDef{})
	if err != nil {
		return nil, err
	}
	colorsData, err := utils.MarshalXMLWithHeader(&diagram.ColorsDef{})
	if err != nil {
		return nil, err
	}

	id := s.getNextShapeID()
	dataPath := fmt.Sprintf("ppt/diagrams/data%d.xml", id)
	layoutPath := fmt.Sprintf("ppt/diagrams/layout%d.xml", id)
	stylePath := fmt.Sprintf("ppt/diagrams/style%d.xml", id)
	colorsPath := fmt.Sprintf("ppt/diagrams/colors%d.xml", id)

	if _, err := s.pres.pkg.AddPart(dataPath, packaging.ContentTypeDiagramData, data); err != nil {
		return nil, err
	}
	if _, err := s.pres.pkg.AddPart(layoutPath, packaging.ContentTypeDiagramLayout, layoutData); err != nil {
		return nil, err
	}
	if _, err := s.pres.pkg.AddPart(stylePath, packaging.ContentTypeDiagramStyle, styleData); err != nil {
		return nil, err
	}
	if _, err := s.pres.pkg.AddPart(colorsPath, packaging.ContentTypeDiagramColors, colorsData); err != nil {
		return nil, err
	}

	sourcePath := s.path
	if sourcePath == "" {
		sourcePath = fmt.Sprintf("ppt/slides/slide%d.xml", s.index+1)
		s.path = sourcePath
	}
	rels := s.pres.pkg.GetRelationships(sourcePath)
	dataRelID := rels.NextID()
	rels.AddWithID(dataRelID, packaging.RelTypeDiagramData, relativeTarget(sourcePath, dataPath), packaging.TargetModeInternal)
	layoutRelID := rels.NextID()
	rels.AddWithID(layoutRelID, packaging.RelTypeDiagramLayout, relativeTarget(sourcePath, layoutPath), packaging.TargetModeInternal)
	styleRelID := rels.NextID()
	rels.AddWithID(styleRelID, packaging.RelTypeDiagramStyle, relativeTarget(sourcePath, stylePath), packaging.TargetModeInternal)
	colorsRelID := rels.NextID()
	rels.AddWithID(colorsRelID, packaging.RelTypeDiagramColors, relativeTarget(sourcePath, colorsPath), packaging.TargetModeInternal)

	name := title
	if name == "" {
		name = fmt.Sprintf("Diagram %d", id)
	}
	gf := &pml.GraphicFrame{
		NvGraphicFramePr: &pml.NvGraphicFramePr{
			CNvPr: &pml.CNvPr{
				ID:   id,
				Name: name,
			},
			CNvGraphicFramePr: &pml.CNvGraphicFramePr{},
			NvPr:              &pml.NvPr{Ph: &pml.Ph{Type: pml.PhTypeDgm}},
		},
		Xfrm: &pml.Xfrm{
			Off: &pml.Off{X: left, Y: top},
			Ext: &pml.Ext{Cx: width, Cy: height},
		},
		Graphic: &dml.Graphic{
			GraphicData: &dml.GraphicData{
				URI: dml.GraphicDataURIDiagram,
				Diagram: &dml.DiagramRef{
					Data:   dataRelID,
					Layout: layoutRelID,
					Colors: colorsRelID,
					Style:  styleRelID,
				},
			},
		},
	}

	s.slide.CSld.SpTree.Content = append(s.slide.CSld.SpTree.Content, gf)
	return &shapeImpl{slide: s, graphicFrame: gf}, nil
}

// Comments returns slide comments.
func (s *slideImpl) Comments() []Comment {
	if s.comments == nil {
		return nil
	}
	result := make([]Comment, len(s.comments.Comment))
	for i, c := range s.comments.Comment {
		result[i] = &commentImpl{slide: s, comment: c}
	}
	return result
}

// Pictures returns all picture shapes on the slide.
func (s *slideImpl) Pictures() []Shape {
	var pictures []Shape
	for _, shape := range s.Shapes() {
		if shape.IsPicture() {
			pictures = append(pictures, shape)
		}
	}
	return pictures
}

// Picture returns a picture shape by name or index (within picture list).
func (s *slideImpl) Picture(identifier string) (Shape, error) {
	pictures := s.Pictures()
	if idx, err := strconv.Atoi(identifier); err == nil {
		if idx < 0 || idx >= len(pictures) {
			return nil, ErrShapeNotFound
		}
		return pictures[idx], nil
	}
	for _, shape := range pictures {
		if shape.Name() == identifier {
			return shape, nil
		}
	}
	return nil, ErrShapeNotFound
}

// ReplacePictureImage replaces the image data for a picture shape.
func (s *slideImpl) ReplacePictureImage(identifier, imagePath string) error {
	pic, err := s.Picture(identifier)
	if err != nil {
		return err
	}
	shape, ok := pic.(*shapeImpl)
	if !ok || shape.pic == nil {
		return ErrShapeNotFound
	}
	if s.pres == nil {
		return ErrInvalidIndex
	}
	target, err := s.pres.addImagePart(imagePath)
	if err != nil {
		return err
	}
	sourcePath := s.path
	if sourcePath == "" {
		sourcePath = fmt.Sprintf("ppt/slides/slide%d.xml", s.index+1)
		s.path = sourcePath
	}
	rels := s.pres.pkg.GetRelationships(sourcePath)
	relID := rels.NextID()
	rels.AddWithID(relID, packaging.RelTypeImage, target, packaging.TargetModeInternal)
	shape.SetImageRelationshipID(relID)
	return nil
}

// AddComment adds a slide comment.
func (s *slideImpl) AddComment(text, author string, x, y float64) (Comment, error) {
	if s == nil || s.pres == nil {
		return nil, ErrInvalidIndex
	}
	authorID := s.pres.ensureCommentAuthor(author)
	if s.comments == nil {
		s.comments = &pml.CommentList{}
	}
	comment := &pml.Comment{
		ID:       newCommentID(),
		AuthorID: authorID,
		Created:  time.Now().Format(time.RFC3339),
		Pos: &pml.CommentPos{
			X: int64(x),
			Y: int64(y),
		},
		TxBody: &pml.CommentText{
			BodyPr:   &pml.CommentBodyPr{},
			LstStyle: &pml.CommentLstStyle{},
			P: []*pml.CommentParagraph{{
				R: []*pml.CommentRun{{
					RPr: &pml.CommentRunProps{Lang: "en-US"},
					T:   text,
				}},
			}},
		},
	}
	s.comments.Comment = append(s.comments.Comment, comment)
	return &commentImpl{slide: s, comment: comment}, nil
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
			if _, ok := item.(*pml.Pic); ok {
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
			if pic, ok := item.(*pml.Pic); ok {
				if pic.NvPicPr != nil && pic.NvPicPr.CNvPr != nil && pic.NvPicPr.CNvPr.Name == identifier {
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

// Layout returns the slide layout if present.
func (s *slideImpl) Layout() SlideLayout {
	if s == nil || s.pres == nil {
		return nil
	}
	if s.path == "" {
		return nil
	}
	rels := s.pres.pkg.GetRelationships(s.path)
	if rels == nil {
		return nil
	}
	rel := rels.FirstByType(packaging.RelTypeSlideLayout)
	if rel == nil {
		return nil
	}
	layoutPath := packaging.ResolveRelationshipTarget(s.path, rel.Target)
	for _, layout := range s.pres.layouts {
		if layout.path == layoutPath {
			return layout
		}
	}
	return &slideLayoutImpl{id: rel.ID, path: layoutPath}
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
				CNvSpPr: &dml.CNvSpPr{SpLocks: &dml.SpLocks{NoGrp: utils.BoolPtr(true)}},
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
		grpXfrm := &pml.GrpXfrm{
			Off:   &pml.Off{X: 0, Y: 0},
			Ext:   &pml.Ext{Cx: 0, Cy: 0},
			ChOff: &pml.Off{X: 0, Y: 0},
			ChExt: &pml.Ext{Cx: 0, Cy: 0},
		}
		s.notes = &pml.Notes{
			CSld: &pml.CSld{
				SpTree: &pml.SpTree{
					NvGrpSpPr: &pml.NvGrpSpPr{
						CNvPr:      &pml.CNvPr{ID: 1, Name: ""},
						CNvGrpSpPr: &pml.CNvGrpSpPr{},
						NvPr:       &pml.NvPr{},
					},
					GrpSpPr: &pml.GrpSpPr{Xfrm: grpXfrm},
				},
			},
			ClrMapOvr: &pml.ClrMapOvr{MasterClrMapping: &pml.MasterClrMapping{}},
		}
	}
}

func newCommentID() string {
	return fmt.Sprintf("{%d}", time.Now().UnixNano())
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
			if pic, ok := item.(*pml.Pic); ok {
				if pic.NvPicPr != nil && pic.NvPicPr.CNvPr != nil && pic.NvPicPr.CNvPr.ID > maxID {
					maxID = pic.NvPicPr.CNvPr.ID
				}
			}
		}
	}
	return maxID + 1
}
