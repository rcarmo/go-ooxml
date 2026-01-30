package presentation

import (
	"errors"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
)

// Errors
var (
	ErrSlideNotFound = errors.New("slide not found")
	ErrShapeNotFound = errors.New("shape not found")
	ErrInvalidIndex  = errors.New("invalid index")
)

// ShapeType represents the type of shape.
type ShapeType int

const (
	ShapeTypeRectangle ShapeType = iota
	ShapeTypeEllipse
	ShapeTypeRoundRect
	ShapeTypeTriangle
	ShapeTypeTextBox
	ShapeTypeLine
	ShapeTypeArrow
)

// PlaceholderType represents the type of placeholder.
type PlaceholderType int

const (
	PlaceholderNone PlaceholderType = iota
	PlaceholderTitle
	PlaceholderCenteredTitle
	PlaceholderSubtitle
	PlaceholderBody
	PlaceholderDate
	PlaceholderFooter
	PlaceholderSlideNumber
	PlaceholderContent
	PlaceholderPicture
	PlaceholderTable
	PlaceholderChart
)

// AutofitType represents text autofit behavior.
type AutofitType int

const (
	AutofitNone   AutofitType = iota // Do not autofit
	AutofitNormal                    // Shrink text to fit
	AutofitShape                     // Resize shape to fit text
)

// BulletType represents bullet point style.
type BulletType int

const (
	BulletNone       BulletType = iota
	BulletAutoNumber            // Numbered list
	BulletCharacter             // Character bullet (•, -, etc.)
	BulletPicture               // Picture bullet
)

// Alignment represents text alignment.
type Alignment int

const (
	AlignmentLeft Alignment = iota
	AlignmentCenter
	AlignmentRight
	AlignmentJustify
)

// =============================================================================
// Type conversion helpers
// =============================================================================

func shapeTypeToGeom(st ShapeType) string {
	switch st {
	case ShapeTypeRectangle:
		return dml.PrstGeomRect
	case ShapeTypeEllipse:
		return dml.PrstGeomEllipse
	case ShapeTypeRoundRect:
		return dml.PrstGeomRoundRect
	case ShapeTypeTriangle:
		return dml.PrstGeomTriangle
	case ShapeTypeTextBox:
		return dml.PrstGeomRect
	case ShapeTypeLine:
		return dml.PrstGeomLine
	case ShapeTypeArrow:
		return dml.PrstGeomRightArrow
	default:
		return dml.PrstGeomRect
	}
}

func geomToShapeType(geom string) ShapeType {
	switch geom {
	case dml.PrstGeomRect:
		return ShapeTypeRectangle
	case dml.PrstGeomEllipse:
		return ShapeTypeEllipse
	case dml.PrstGeomRoundRect:
		return ShapeTypeRoundRect
	case dml.PrstGeomTriangle:
		return ShapeTypeTriangle
	case dml.PrstGeomLine:
		return ShapeTypeLine
	case dml.PrstGeomRightArrow, dml.PrstGeomLeftArrow, dml.PrstGeomUpArrow, dml.PrstGeomDownArrow:
		return ShapeTypeArrow
	default:
		return ShapeTypeRectangle
	}
}

func phTypeToPlaceholderType(phType string) PlaceholderType {
	switch phType {
	case pml.PhTypeTitle:
		return PlaceholderTitle
	case pml.PhTypeCtrTitle:
		return PlaceholderCenteredTitle
	case pml.PhTypeSubTitle:
		return PlaceholderSubtitle
	case pml.PhTypeBody:
		return PlaceholderBody
	case pml.PhTypeDT:
		return PlaceholderDate
	case pml.PhTypeFtr:
		return PlaceholderFooter
	case pml.PhTypeSldNum:
		return PlaceholderSlideNumber
	case pml.PhTypeTbl:
		return PlaceholderTable
	case pml.PhTypeChart:
		return PlaceholderChart
	case pml.PhTypePic:
		return PlaceholderPicture
	default:
		return PlaceholderContent
	}
}

func placeholderTypeToPhType(pt PlaceholderType) string {
	switch pt {
	case PlaceholderTitle:
		return pml.PhTypeTitle
	case PlaceholderCenteredTitle:
		return pml.PhTypeCtrTitle
	case PlaceholderSubtitle:
		return pml.PhTypeSubTitle
	case PlaceholderBody:
		return pml.PhTypeBody
	case PlaceholderDate:
		return pml.PhTypeDT
	case PlaceholderFooter:
		return pml.PhTypeFtr
	case PlaceholderSlideNumber:
		return pml.PhTypeSldNum
	case PlaceholderTable:
		return pml.PhTypeTbl
	case PlaceholderChart:
		return pml.PhTypeChart
	case PlaceholderPicture:
		return pml.PhTypePic
	default:
		return ""
	}
}
