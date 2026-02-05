package presentation

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
)

// Shape represents a shape on a slide.
type shapeImpl struct {
	slide *slideImpl
	sp    *dml.Sp
	graphicFrame *pml.GraphicFrame
	pic   *pml.Pic
}

// ID returns the shape ID.
func (s *shapeImpl) ID() int {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.CNvPr != nil {
		return s.sp.NvSpPr.CNvPr.ID
	}
	if s.graphicFrame != nil && s.graphicFrame.NvGraphicFramePr != nil && s.graphicFrame.NvGraphicFramePr.CNvPr != nil {
		return s.graphicFrame.NvGraphicFramePr.CNvPr.ID
	}
	if s.pic != nil && s.pic.NvPicPr != nil && s.pic.NvPicPr.CNvPr != nil {
		return s.pic.NvPicPr.CNvPr.ID
	}
	return 0
}

// Name returns the shape name.
func (s *shapeImpl) Name() string {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.CNvPr != nil {
		return s.sp.NvSpPr.CNvPr.Name
	}
	if s.graphicFrame != nil && s.graphicFrame.NvGraphicFramePr != nil && s.graphicFrame.NvGraphicFramePr.CNvPr != nil {
		return s.graphicFrame.NvGraphicFramePr.CNvPr.Name
	}
	if s.pic != nil && s.pic.NvPicPr != nil && s.pic.NvPicPr.CNvPr != nil {
		return s.pic.NvPicPr.CNvPr.Name
	}
	return ""
}

// SetName sets the shape name.
func (s *shapeImpl) SetName(name string) {
	if s.graphicFrame != nil {
		s.ensureNvGraphicFramePr()
		s.graphicFrame.NvGraphicFramePr.CNvPr.Name = name
		return
	}
	if s.pic != nil {
		if s.pic.NvPicPr == nil {
			s.pic.NvPicPr = &pml.NvPicPr{}
		}
		if s.pic.NvPicPr.CNvPr == nil {
			s.pic.NvPicPr.CNvPr = &pml.CNvPr{}
		}
		s.pic.NvPicPr.CNvPr.Name = name
		return
	}
	s.ensureNvSpPr()
	s.sp.NvSpPr.CNvPr.Name = name
}

// Type returns the shape type.
func (s *shapeImpl) Type() ShapeType {
	if s.graphicFrame != nil {
		if s.graphicFrame.Graphic != nil && s.graphicFrame.Graphic.GraphicData != nil && s.graphicFrame.Graphic.GraphicData.Tbl != nil {
			return ShapeTypeTable
		}
		return ShapeTypeRectangle
	}
	if s.pic != nil {
		return ShapeTypeRectangle
	}
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.CNvSpPr != nil && s.sp.NvSpPr.CNvSpPr.TxBox != nil && *s.sp.NvSpPr.CNvSpPr.TxBox {
		return ShapeTypeTextBox
	}
	if s.sp.SpPr != nil && s.sp.SpPr.PrstGeom != nil {
		return geomToShapeType(s.sp.SpPr.PrstGeom.Prst)
	}
	return ShapeTypeRectangle
}

// IsPicture returns true if this shape is a picture.
func (s *shapeImpl) IsPicture() bool {
	return s.pic != nil
}

// ImageRelationshipID returns the image relationship ID for a picture shape.
func (s *shapeImpl) ImageRelationshipID() string {
	if s.pic == nil || s.pic.BlipFill == nil || s.pic.BlipFill.Blip == nil {
		return ""
	}
	return s.pic.BlipFill.Blip.Embed
}

// SetImageRelationshipID sets the image relationship ID for a picture shape.
func (s *shapeImpl) SetImageRelationshipID(relID string) {
	if s.pic == nil {
		return
	}
	if s.pic.BlipFill == nil {
		s.pic.BlipFill = &dml.BlipFill{}
	}
	if s.pic.BlipFill.Blip == nil {
		s.pic.BlipFill.Blip = &dml.Blip{}
	}
	s.pic.BlipFill.Blip.Embed = relID
}

// HasTable returns true if the shape contains a table.
func (s *shapeImpl) HasTable() bool {
	return s.graphicFrame != nil && s.graphicFrame.Graphic != nil && s.graphicFrame.Graphic.GraphicData != nil && s.graphicFrame.Graphic.GraphicData.Tbl != nil
}

// Table returns the table if present.
func (s *shapeImpl) Table() Table {
	if !s.HasTable() {
		return nil
	}
	return tableFromGraphicFrame(s.graphicFrame)
}

// IsPlaceholder returns true if this is a placeholder shape.
func (s *shapeImpl) IsPlaceholder() bool {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.NvPr != nil && s.sp.NvSpPr.NvPr.Ph != nil {
		return true
	}
	if s.graphicFrame != nil && s.graphicFrame.NvGraphicFramePr != nil && s.graphicFrame.NvGraphicFramePr.NvPr != nil && s.graphicFrame.NvGraphicFramePr.NvPr.Ph != nil {
		return true
	}
	if s.pic != nil && s.pic.NvPicPr != nil && s.pic.NvPicPr.NvPr != nil && s.pic.NvPicPr.NvPr.Ph != nil {
		return true
	}
	return false
}

// PlaceholderType returns the placeholder type.
func (s *shapeImpl) PlaceholderType() PlaceholderType {
	if s.sp.NvSpPr != nil && s.sp.NvSpPr.NvPr != nil && s.sp.NvSpPr.NvPr.Ph != nil {
		return phTypeToPlaceholderType(s.sp.NvSpPr.NvPr.Ph.Type)
	}
	if s.graphicFrame != nil && s.graphicFrame.NvGraphicFramePr != nil && s.graphicFrame.NvGraphicFramePr.NvPr != nil && s.graphicFrame.NvGraphicFramePr.NvPr.Ph != nil {
		return phTypeToPlaceholderType(s.graphicFrame.NvGraphicFramePr.NvPr.Ph.Type)
	}
	if s.pic != nil && s.pic.NvPicPr != nil && s.pic.NvPicPr.NvPr != nil && s.pic.NvPicPr.NvPr.Ph != nil {
		return phTypeToPlaceholderType(s.pic.NvPicPr.NvPr.Ph.Type)
	}
	return PlaceholderNone
}

// =============================================================================
// Position and Size
// =============================================================================

// Left returns the left position in EMUs.
func (s *shapeImpl) Left() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Off != nil {
		return s.sp.SpPr.Xfrm.Off.X
	}
	if s.graphicFrame != nil && s.graphicFrame.Xfrm != nil && s.graphicFrame.Xfrm.Off != nil {
		return s.graphicFrame.Xfrm.Off.X
	}
	if s.pic != nil && s.pic.SpPr != nil && s.pic.SpPr.Xfrm != nil && s.pic.SpPr.Xfrm.Off != nil {
		return s.pic.SpPr.Xfrm.Off.X
	}
	return 0
}

// Top returns the top position in EMUs.
func (s *shapeImpl) Top() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Off != nil {
		return s.sp.SpPr.Xfrm.Off.Y
	}
	if s.graphicFrame != nil && s.graphicFrame.Xfrm != nil && s.graphicFrame.Xfrm.Off != nil {
		return s.graphicFrame.Xfrm.Off.Y
	}
	if s.pic != nil && s.pic.SpPr != nil && s.pic.SpPr.Xfrm != nil && s.pic.SpPr.Xfrm.Off != nil {
		return s.pic.SpPr.Xfrm.Off.Y
	}
	return 0
}

// Width returns the width in EMUs.
func (s *shapeImpl) Width() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Ext != nil {
		return s.sp.SpPr.Xfrm.Ext.Cx
	}
	if s.graphicFrame != nil && s.graphicFrame.Xfrm != nil && s.graphicFrame.Xfrm.Ext != nil {
		return s.graphicFrame.Xfrm.Ext.Cx
	}
	if s.pic != nil && s.pic.SpPr != nil && s.pic.SpPr.Xfrm != nil && s.pic.SpPr.Xfrm.Ext != nil {
		return s.pic.SpPr.Xfrm.Ext.Cx
	}
	return 0
}

// Height returns the height in EMUs.
func (s *shapeImpl) Height() int64 {
	if s.sp.SpPr != nil && s.sp.SpPr.Xfrm != nil && s.sp.SpPr.Xfrm.Ext != nil {
		return s.sp.SpPr.Xfrm.Ext.Cy
	}
	if s.graphicFrame != nil && s.graphicFrame.Xfrm != nil && s.graphicFrame.Xfrm.Ext != nil {
		return s.graphicFrame.Xfrm.Ext.Cy
	}
	if s.pic != nil && s.pic.SpPr != nil && s.pic.SpPr.Xfrm != nil && s.pic.SpPr.Xfrm.Ext != nil {
		return s.pic.SpPr.Xfrm.Ext.Cy
	}
	return 0
}

// SetPosition sets the position in EMUs.
func (s *shapeImpl) SetPosition(left, top int64) {
	if s.graphicFrame != nil {
		s.ensureGraphicFrameXfrm()
		if s.graphicFrame.Xfrm.Off == nil {
			s.graphicFrame.Xfrm.Off = &pml.Off{}
		}
		s.graphicFrame.Xfrm.Off.X = left
		s.graphicFrame.Xfrm.Off.Y = top
		return
	}
	if s.pic != nil {
		if s.pic.SpPr == nil {
			s.pic.SpPr = &dml.SpPr{}
		}
		if s.pic.SpPr.Xfrm == nil {
			s.pic.SpPr.Xfrm = &dml.Xfrm{}
		}
		if s.pic.SpPr.Xfrm.Off == nil {
			s.pic.SpPr.Xfrm.Off = &dml.Off{}
		}
		s.pic.SpPr.Xfrm.Off.X = left
		s.pic.SpPr.Xfrm.Off.Y = top
		return
	}
	s.ensureXfrm()
	if s.sp.SpPr.Xfrm.Off == nil {
		s.sp.SpPr.Xfrm.Off = &dml.Off{}
	}
	s.sp.SpPr.Xfrm.Off.X = left
	s.sp.SpPr.Xfrm.Off.Y = top
}

// SetSize sets the size in EMUs.
func (s *shapeImpl) SetSize(width, height int64) {
	if s.graphicFrame != nil {
		s.ensureGraphicFrameXfrm()
		if s.graphicFrame.Xfrm.Ext == nil {
			s.graphicFrame.Xfrm.Ext = &pml.Ext{}
		}
		s.graphicFrame.Xfrm.Ext.Cx = width
		s.graphicFrame.Xfrm.Ext.Cy = height
		return
	}
	if s.pic != nil {
		if s.pic.SpPr == nil {
			s.pic.SpPr = &dml.SpPr{}
		}
		if s.pic.SpPr.Xfrm == nil {
			s.pic.SpPr.Xfrm = &dml.Xfrm{}
		}
		if s.pic.SpPr.Xfrm.Ext == nil {
			s.pic.SpPr.Xfrm.Ext = &dml.Ext{}
		}
		s.pic.SpPr.Xfrm.Ext.Cx = width
		s.pic.SpPr.Xfrm.Ext.Cy = height
		return
	}
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
func (s *shapeImpl) HasTextFrame() bool {
	if s.graphicFrame != nil {
		return false
	}
	if s.pic != nil {
		return false
	}
	if s.sp == nil {
		return false
	}
	return s.sp.TxBody != nil
}

// TextFrame returns the text frame for the shape.
func (s *shapeImpl) TextFrame() TextFrame {
	if s.graphicFrame != nil {
		return nil
	}
	if s.pic != nil {
		return nil
	}
	if s.sp == nil {
		return nil
	}
	if s.sp.TxBody == nil {
		s.sp.TxBody = &dml.TxBody{
			BodyPr:   &dml.BodyPr{},
			LstStyle: &dml.LstStyle{},
			P:        []*dml.P{{}},
		}
	}
	return &textFrameImpl{
		shape:  s,
		txBody: s.sp.TxBody,
	}
}

// Text returns all text in the shape as a single string.
func (s *shapeImpl) Text() string {
	if tf, ok := s.TextFrame().(*textFrameImpl); ok {
		return tf.Text()
	}
	return ""
}

// SetText sets the text of the shape, replacing all existing content.
func (s *shapeImpl) SetText(text string) error {
	if s.graphicFrame != nil {
		return ErrShapeNotFound
	}
	if s.pic != nil {
		return ErrShapeNotFound
	}
	if tf, ok := s.TextFrame().(*textFrameImpl); ok {
		tf.SetText(text)
		return nil
	}
	return ErrShapeNotFound
}

// =============================================================================
// Fill and Outline
// =============================================================================

// SetFillColor sets a solid fill color (hex format like "FF0000").
func (s *shapeImpl) SetFillColor(hex string) {
	if s.sp == nil {
		return
	}
	if s.pic != nil {
		return
	}
	s.ensureSpPr()
	s.sp.SpPr.NoFill = nil
	s.sp.SpPr.GradFill = nil
	s.sp.SpPr.BlipFill = nil
	s.sp.SpPr.SolidFill = &dml.SolidFill{
		SrgbClr: &dml.SrgbClr{Val: hex},
	}
}

// SetNoFill removes the fill from the shape.
func (s *shapeImpl) SetNoFill() {
	if s.sp == nil {
		return
	}
	if s.pic != nil {
		return
	}
	s.ensureSpPr()
	s.sp.SpPr.SolidFill = nil
	s.sp.SpPr.GradFill = nil
	s.sp.SpPr.BlipFill = nil
	s.sp.SpPr.NoFill = &dml.NoFill{}
}

// SetLineColor sets the outline color (hex format).
func (s *shapeImpl) SetLineColor(hex string, widthEMU int64) {
	if s.sp == nil {
		return
	}
	if s.pic != nil {
		return
	}
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

func (s *shapeImpl) ensureNvSpPr() {
	if s.sp == nil {
		return
	}
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

func (s *shapeImpl) ensureSpPr() {
	if s.sp == nil {
		return
	}
	if s.sp.SpPr == nil {
		s.sp.SpPr = &dml.SpPr{}
	}
}

func (s *shapeImpl) ensureXfrm() {
	if s.sp == nil {
		return
	}
	s.ensureSpPr()
	if s.sp.SpPr.Xfrm == nil {
		s.sp.SpPr.Xfrm = &dml.Xfrm{}
	}
}

func (s *shapeImpl) ensureNvGraphicFramePr() {
	if s.graphicFrame.NvGraphicFramePr == nil {
		s.graphicFrame.NvGraphicFramePr = &pml.NvGraphicFramePr{}
	}
	if s.graphicFrame.NvGraphicFramePr.CNvPr == nil {
		s.graphicFrame.NvGraphicFramePr.CNvPr = &pml.CNvPr{ID: 1}
	}
	if s.graphicFrame.NvGraphicFramePr.CNvGraphicFramePr == nil {
		s.graphicFrame.NvGraphicFramePr.CNvGraphicFramePr = &pml.CNvGraphicFramePr{}
	}
	if s.graphicFrame.NvGraphicFramePr.NvPr == nil {
		s.graphicFrame.NvGraphicFramePr.NvPr = &pml.NvPr{}
	}
}

func (s *shapeImpl) ensureGraphicFrameXfrm() {
	if s.graphicFrame.Xfrm == nil {
		s.graphicFrame.Xfrm = &pml.Xfrm{}
	}
}
