package presentation

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
)

// Shape represents a shape on a slide.
type Shape struct {
	slide *Slide
	sp    *dml.Sp
}

// ID returns the shape ID.
func (s *Shape) ID() int {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.CNvPr != nil {
		return s.sp.NvSpPr.CNvPr.ID
	}
	return 0
}

// Name returns the shape name.
func (s *Shape) Name() string {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.CNvPr != nil {
		return s.sp.NvSpPr.CNvPr.Name
	}
	return ""
}

// SetName sets the shape name.
func (s *Shape) SetName(name string) {
	s.ensureNvSpPr()
	s.sp.NvSpPr.CNvPr.Name = name
}

// Type returns the shape type.
func (s *Shape) Type() ShapeType {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.CNvSpPr != nil && s.sp.NvSpPr.CNvSpPr.TxBox != nil && *s.sp.NvSpPr.CNvSpPr.TxBox {
		return ShapeTypeTextBox
	}
	if s.sp.SpPr != nil && s.sp.SpPr.PrstGeom != nil {
		return geomToShapeType(s.sp.SpPr.PrstGeom.Prst)
	}
	return ShapeTypeRectangle
}

// IsPlaceholder returns true if this is a placeholder shape.
func (s *Shape) IsPlaceholder() bool {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.NvPr != nil && s.sp.NvSpPr.NvPr.Ph != nil {
		return true
	}
	return false
}

// PlaceholderType returns the placeholder type.
func (s *Shape) PlaceholderType() PlaceholderType {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.NvPr != nil && s.sp.NvSpPr.NvPr.Ph != nil {
		return phTypeToPlaceholderType(s.sp.NvSpPr.NvPr.Ph.Type)
	}
	return PlaceholderNone
}

// =============================================================================
// Position and Size
// =============================================================================

// Left returns the left position in EMUs.
func (s *Shape) Left() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Off != nil {
		return s.sp.SpPr.Xfrm.Off.X
	}
	return 0
}

// Top returns the top position in EMUs.
func (s *Shape) Top() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Off != nil {
		return s.sp.SpPr.Xfrm.Off.Y
	}
	return 0
}

// Width returns the width in EMUs.
func (s *Shape) Width() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Ext != nil {
		return s.sp.SpPr.Xfrm.Ext.Cx
	}
	return 0
}

// Height returns the height in EMUs.
func (s *Shape) Height() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Ext != nil {
		return s.sp.SpPr.Xfrm.Ext.Cy
	}
	return 0
}

// SetPosition sets the position in EMUs.
func (s *Shape) SetPosition(left, top int64) {
	s.ensureXfrm()
	if s.sp.SpPr.Xfrm.Off == nil {
		s.sp.SpPr.Xfrm.Off = &dml.Off{}
	}
	s.sp.SpPr.Xfrm.Off.X = left
	s.sp.SpPr.Xfrm.Off.Y = top
}

// SetSize sets the size in EMUs.
func (s *Shape) SetSize(width, height int64) {
	s.ensureXfrm()
	if s.sp.SpPr.Xfrm.Ext == nil {
		s.sp.SpPr.Xfrm.Ext = &dml.Ext{}
	}
	s.sp.SpPr.Xfrm.Ext.Cx = width
	s.sp.SpPr.Xfrm.Ext.Cy = height
}

// =============================================================================
// Text content
// =============================================================================

// HasTextFrame returns true if the shape has a text frame.
func (s *Shape) HasTextFrame() bool {
	return s.sp.TxBody != nil
}

// TextFrame returns the text frame for the shape.
func (s *Shape) TextFrame() *TextFrame {
	if s.sp.TxBody == nil {
		s.sp.TxBody = &dml.TxBody{
			BodyPr:   &dml.BodyPr{},
			LstStyle: &dml.LstStyle{},
			P:        []*dml.P{{}},
		}
	}
	return &TextFrame{
		shape:  s,
		txBody: s.sp.TxBody,
	}
}

// Text returns all text in the shape as a single string.
func (s *Shape) Text() string {
	if s.sp.TxBody == nil {
		return ""
	}

	var paragraphs []string
	for _, p := range s.sp.TxBody.P {
		var runs []string
		for _, r := range p.R {
			runs = append(runs, r.T)
		}
		paragraphs = append(paragraphs, strings.Join(runs, ""))
	}
	return strings.Join(paragraphs, "\n")
}

// SetText sets the text of the shape, replacing all existing content.
func (s *Shape) SetText(text string) error {
	if s.sp.TxBody == nil {
		s.sp.TxBody = &dml.TxBody{
			BodyPr:   &dml.BodyPr{},
			LstStyle: &dml.LstStyle{},
		}
	}

	// Split text into paragraphs
	lines := strings.Split(text, "\n")
	s.sp.TxBody.P = make([]*dml.P, len(lines))
	for i, line := range lines {
		s.sp.TxBody.P[i] = &dml.P{
			R: []*dml.R{{T: line}},
		}
	}

	return nil
}

// =============================================================================
// Fill and Outline
// =============================================================================

// SetFillColor sets a solid fill color (hex format like "FF0000").
func (s *Shape) SetFillColor(hex string) {
	s.ensureSpPr()
	s.sp.SpPr.NoFill = nil
	s.sp.SpPr.GradFill = nil
	s.sp.SpPr.BlipFill = nil
	s.sp.SpPr.SolidFill = &dml.SolidFill{
		SrgbClr: &dml.SrgbClr{Val: hex},
	}
}

// SetNoFill removes the fill from the shape.
func (s *Shape) SetNoFill() {
	s.ensureSpPr()
	s.sp.SpPr.SolidFill = nil
	s.sp.SpPr.GradFill = nil
	s.sp.SpPr.BlipFill = nil
	s.sp.SpPr.NoFill = &dml.NoFill{}
}

// SetLineColor sets the outline color (hex format).
func (s *Shape) SetLineColor(hex string, widthEMU int64) {
	s.ensureSpPr()
	if s.sp.SpPr.Ln == nil {
		s.sp.SpPr.Ln = &dml.Ln{}
	}
	s.sp.SpPr.Ln.W = widthEMU
	s.sp.SpPr.Ln.SolidFill = &dml.SolidFill{
		SrgbClr: &dml.SrgbClr{Val: hex},
	}
}

// =============================================================================
// Internal methods
// =============================================================================

func (s *Shape) ensureNvSpPr() {
	if s.sp.NvSpPr == nil {
		s.sp.NvSpPr = &dml.NvSpPr{}
	}
	if s.sp.NvSpPr.CNvPr == nil {
		s.sp.NvSpPr.CNvPr = &dml.CNvPr{ID: 1}
	}
	if s.sp.NvSpPr.CNvSpPr == nil {
		s.sp.NvSpPr.CNvSpPr = &dml.CNvSpPr{}
	}
}

func (s *Shape) ensureSpPr() {
	if s.sp.SpPr == nil {
		s.sp.SpPr = &dml.SpPr{}
	}
}

func (s *Shape) ensureXfrm() {
	s.ensureSpPr()
	if s.sp.SpPr.Xfrm == nil {
		s.sp.SpPr.Xfrm = &dml.Xfrm{}
	}
}
