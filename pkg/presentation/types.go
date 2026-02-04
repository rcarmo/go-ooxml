package presentation

import (
	"errors"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
)

var (
	// ErrSlideNotFound is returned when a slide index is invalid.
	ErrSlideNotFound = errors.New("slide not found")
	// ErrShapeNotFound is returned when a slide shape cannot be located.
	ErrShapeNotFound = errors.New("shape not found")
	// ErrInvalidIndex is returned when an index is out of range.
	ErrInvalidIndex  = errors.New("invalid index")
)

// PresentationProperties maps to core properties.
type PresentationProperties = common.CoreProperties

// Presentation represents a PowerPoint presentation.
type Presentation interface {
	Save() error
	SaveAs(path string) error
	Close() error
	Slides() []Slide
	Slide(index int) (Slide, error)
	AddSlide(layoutIndex int) Slide
	InsertSlide(index, layoutIndex int) Slide
	DeleteSlide(index int) error
	DuplicateSlide(index int) Slide
	ReorderSlides(newOrder []int) error
	SlidesRaw() []*slideImpl
	SlideCount() int
	CoreProperties() (*common.CoreProperties, error)
	SetCoreProperties(props *common.CoreProperties) error
	Masters() []SlideMaster
	Layouts() []SlideLayout
	Properties() PresentationProperties
	SlideSize() (width, height int64)
	SetSlideSize(width, height int64) error
}

// Slide represents a slide in the presentation.
type Slide interface {
	Index() int
	ID() string
	Hidden() bool
	SetHidden(v bool)
	Layout() SlideLayout
	Shapes() []Shape
	Shape(identifier string) (Shape, error)
	AddShape(shapeType ShapeType) Shape
	AddTextBox(left, top, width, height int64) Shape
	AddTable(rows, cols int, left, top, width, height int64) Table
	AddPicture(imagePath string, left, top, width, height int64) (Shape, error)
	DeleteShape(identifier string) error
	Placeholders() []Shape
	TitlePlaceholder() Shape
	BodyPlaceholder() Shape
	Tables() []Table
	Notes() string
	SetNotes(text string) error
	AppendNotes(text string) error
	HasNotes() bool
	Comments() []Comment
	AddComment(text, author string, x, y float64) (Comment, error)
}

// Shape represents a shape on a slide.
type Shape interface {
	ID() int
	Name() string
	SetName(name string)
	Type() ShapeType
	IsPlaceholder() bool
	PlaceholderType() PlaceholderType
	Left() int64
	Top() int64
	Width() int64
	Height() int64
	SetPosition(left, top int64)
	SetSize(width, height int64)
	HasTextFrame() bool
	TextFrame() TextFrame
	Text() string
	SetText(text string) error
	HasTable() bool
	Table() Table
	SetFillColor(hex string)
	SetNoFill()
	SetLineColor(hex string, widthEMU int64)
}

// TextFrame represents the text content of a shape.
type TextFrame interface {
	Text() string
	SetText(text string)
	Paragraphs() []TextParagraph
	AddParagraph() TextParagraph
	ClearParagraphs()
	AutofitType() AutofitType
	SetAutofitType(t AutofitType)
}

// TextParagraph represents a paragraph in a text frame.
type TextParagraph interface {
	Text() string
	SetText(text string)
	Runs() []TextRun
	AddRun() TextRun
	BulletType() BulletType
	SetBulletType(t BulletType)
	Level() int
	SetLevel(level int)
	Alignment() Alignment
	SetAlignment(a Alignment)
}

// TextRun represents a text run in a paragraph.
type TextRun interface {
	Text() string
	SetText(text string)
	Bold() bool
	SetBold(v bool)
	Italic() bool
	SetItalic(v bool)
	Underline() bool
	SetUnderline(v bool)
	FontSize() float64
	SetFontSize(points float64)
	FontName() string
	SetFontName(name string)
	Color() string
	SetColor(hex string)
}

// Table represents a table in a slide.
type Table interface {
	Rows() []TableRow
	Row(index int) TableRow
	Cell(row, col int) TableCell
	AddRow() TableRow
	InsertRow(index int) TableRow
	DeleteRow(index int) error
	RowCount() int
	ColumnCount() int
}

// TableRow represents a table row.
type TableRow interface {
	Cells() []TableCell
	Cell(index int) TableCell
	Height() int64
	SetHeight(height int64)
}

// TableCell represents a table cell.
type TableCell interface {
	TextFrame() TextFrame
	Text() string
	SetText(text string)
	RowSpan() int
	ColSpan() int
	SetRowSpan(span int)
	SetColSpan(span int)
}

// SlideMaster represents a slide master.
type SlideMaster interface {
	ID() string
	Path() string
}

// SlideLayout represents a slide layout.
type SlideLayout interface {
	ID() string
	Path() string
}

// Comment represents a slide comment.
type Comment interface {
	ID() string
	Author() string
	Text() string
	SetText(text string)
}

// ShapeType represents the type of shape.
type ShapeType int

const (
	// ShapeTypeRectangle identifies a rectangle shape.
	ShapeTypeRectangle ShapeType = iota
	// ShapeTypeEllipse identifies an ellipse shape.
	ShapeTypeEllipse
	// ShapeTypeRoundRect identifies a rounded rectangle shape.
	ShapeTypeRoundRect
	// ShapeTypeTriangle identifies a triangle shape.
	ShapeTypeTriangle
	// ShapeTypeTextBox identifies a text box shape.
	ShapeTypeTextBox
	// ShapeTypeTable identifies a table shape.
	ShapeTypeTable
	// ShapeTypeLine identifies a line shape.
	ShapeTypeLine
	// ShapeTypeArrow identifies an arrow shape.
	ShapeTypeArrow
)

// PlaceholderType represents the type of placeholder.
type PlaceholderType int

const (
	// PlaceholderNone indicates a non-placeholder shape.
	PlaceholderNone PlaceholderType = iota
	// PlaceholderTitle indicates a title placeholder.
	PlaceholderTitle
	// PlaceholderCenteredTitle indicates a centered title placeholder.
	PlaceholderCenteredTitle
	// PlaceholderSubtitle indicates a subtitle placeholder.
	PlaceholderSubtitle
	// PlaceholderBody indicates a body placeholder.
	PlaceholderBody
	// PlaceholderDate indicates a date placeholder.
	PlaceholderDate
	// PlaceholderFooter indicates a footer placeholder.
	PlaceholderFooter
	// PlaceholderSlideNumber indicates a slide number placeholder.
	PlaceholderSlideNumber
	// PlaceholderContent indicates a content placeholder.
	PlaceholderContent
	// PlaceholderPicture indicates a picture placeholder.
	PlaceholderPicture
	// PlaceholderTable indicates a table placeholder.
	PlaceholderTable
	// PlaceholderChart indicates a chart placeholder.
	PlaceholderChart
)

// AutofitType represents text autofit behavior.
type AutofitType int

const (
	// AutofitNone disables text autofit.
	AutofitNone AutofitType = iota
	// AutofitNormal shrinks text to fit the shape.
	AutofitNormal
	// AutofitShape resizes the shape to fit text.
	AutofitShape
)

// BulletType represents bullet point style.
type BulletType int

const (
	// BulletNone disables bullets.
	BulletNone BulletType = iota
	// BulletAutoNumber uses numbered bullets.
	BulletAutoNumber
	// BulletCharacter uses a character bullet (•, -, etc.).
	BulletCharacter
	// BulletPicture uses a picture bullet.
	BulletPicture
)

// Alignment represents text alignment.
type Alignment int

const (
	// AlignmentLeft aligns text to the left.
	AlignmentLeft Alignment = iota
	// AlignmentCenter centers the text.
	AlignmentCenter
	// AlignmentRight aligns text to the right.
	AlignmentRight
	// AlignmentJustify justifies the text.
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
	case ShapeTypeTable:
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
