package document

import (
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// DocumentProperties maps to core properties.
type DocumentProperties = common.CoreProperties

// ParagraphProperties represents paragraph properties.
type ParagraphProperties = wml.PPr

// RunProperties represents run properties.
type RunProperties = wml.RPr

// VerticalMerge represents the vertical merge state of a cell.
type VerticalMerge string

// BodyElement represents a body element.
type BodyElement interface {
	Index() int
}

// RevisionLocation is a placeholder for revision location.
type RevisionLocation struct{}

// Revision represents a tracked change.
type Revision interface {
	ID() string
	Type() RevisionType
	Author() string
	Date() time.Time
	Text() string
	Location() RevisionLocation
	Accept() error
	Reject() error
}

// TrackChanges provides track change functionality.
type TrackChanges interface {
	Enabled() bool
	Enable()
	Disable()
	Author() string
	SetAuthor(name string)
	Insertions() []Revision
	Deletions() []Revision
	AllRevisions() []Revision
	AcceptAll()
	RejectAll()
	AcceptRevision(id string) error
	RejectRevision(id string) error
	InsertText(para Paragraph, position int, text string) error
	DeleteText(para Paragraph, start, end int) error
	ReplaceText(para Paragraph, oldText, newText string) error
}

// Comments provides comment functionality.
type Comments interface {
	All() []Comment
	ByID(id string) (Comment, error)
	Add(text, author string, anchorText string) (Comment, error)
	Delete(id string) error
}

// Comment represents a document comment.
type Comment interface {
	ID() string
	Author() string
	SetAuthor(author string)
	Initials() string
	SetInitials(initials string)
	Date() time.Time
	Text() string
	SetText(text string)
	AnchoredText() string
	Replies() []Comment
	AddReply(text, author string) (Comment, error)
}

// Styles represents document style management.
type Styles interface {
	All() []Style
	ByID(id string) Style
	ByName(name string) Style
	AddParagraphStyle(id, name string) Style
	AddCharacterStyle(id, name string) Style
	AddTableStyle(id, name string) Style
	AddNumberingStyle(id, name string) Style
	Delete(id string) bool
	DefaultParagraphStyle() Style
	DefaultCharacterStyle() Style
	List() []Style
}

// Style represents a document style.
type Style interface {
	ID() string
	Name() string
	SetName(name string)
	Type() StyleType
	BasedOn() string
	SetBasedOn(styleID string)
	IsDefault() bool
	SetDefault(v bool)
	SetBold(v bool)
	SetItalic(v bool)
	SetFontSize(points float64)
	SetFontName(name string)
	SetColor(hex string)
	SetAlignment(align string)
	SetSpacingBefore(twips int64)
	SetSpacingAfter(twips int64)
	ParagraphProperties() *ParagraphProperties
	SetParagraphProperties(ppr *ParagraphProperties)
	RunProperties() *RunProperties
	SetRunProperties(rpr *RunProperties)
}

// Header represents a document header.
type Header interface {
	Type() HeaderFooterType
	Paragraphs() []Paragraph
	AddParagraph() Paragraph
	Text() string
	SetText(text string)
}

// Footer represents a document footer.
type Footer interface {
	Type() HeaderFooterType
	Paragraphs() []Paragraph
	AddParagraph() Paragraph
	Text() string
	SetText(text string)
}

// Document represents a Word document.
type Document interface {
	Save() error
	SaveAs(path string) error
	Close() error
	Body() Body
	Paragraphs() []Paragraph
	Tables() []Table
	Sections() []Section
	AddParagraph() Paragraph
	AddTable(rows, cols int) Table
	AddParagraphStyle(id, name string) Style
	AddCharacterStyle(id, name string) Style
	AddTableStyle(id, name string) Style
	AddNumberingStyle(id, name string) Style
	DefaultParagraphStyle() Style
	DefaultCharacterStyle() Style
	Styles() Styles
	TrackChanges() TrackChanges
	TrackChangesEnabled() bool
	EnableTrackChanges(author string)
	DisableTrackChanges()
	TrackAuthor() string
	SetTrackAuthor(author string)
	AllRevisions() []Revision
	Insertions() []Revision
	Deletions() []Revision
	AcceptAllRevisions()
	RejectAllRevisions()
	Comments() Comments
	CommentByID(id string) Comment
	DeleteComment(id string) error
	CoreProperties() (*common.CoreProperties, error)
	SetCoreProperties(props *common.CoreProperties) error
	Headers() []Header
	AddHeader(hfType HeaderFooterType) Header
	Footers() []Footer
	AddFooter(hfType HeaderFooterType) Footer
	Header(hfType HeaderFooterType) Header
	Footer(hfType HeaderFooterType) Footer
	StyleByID(id string) Style
	StyleByName(name string) Style
	DeleteStyle(id string) bool
	AddNumberedListStyle() (int, error)
	Numbering() []*Numbering
	ContentControlsByTag(tag string) []*ContentControl
	Properties() DocumentProperties
	ContentControls() []*ContentControl
	AddContentControl(tag, alias, text string) *ContentControl
	AddBlockContentControl(tag, alias, text string) *ContentControl
	ContentControlByTag(tag string) *ContentControl
}


// Body represents the document body.
type Body interface {
	Elements() []BodyElement
	Paragraphs() []Paragraph
	Tables() []Table
	ContentControls() []*ContentControl
	AddParagraph() Paragraph
	AddTable(rows, cols int) Table
	InsertParagraphAt(index int) Paragraph
	InsertParagraphBefore(target BodyElement) Paragraph
	InsertParagraphAfter(target BodyElement) Paragraph
	ElementCount() int
}

// Section represents a document section.
type Section interface {
	Header(hfType HeaderFooterType) Header
	Footer(hfType HeaderFooterType) Footer
	AddHeader(hfType HeaderFooterType) Header
	AddFooter(hfType HeaderFooterType) Footer
}
// Paragraph represents a paragraph.
type Paragraph interface {
	BodyElement
	Text() string
	SetText(text string)
	Runs() []Run
	AddRun() Run
	InsertTrackedText(text string) Run
	DeleteTrackedText(runIndex int) error
	Style() string
	SetStyle(styleID string)
	Properties() ParagraphProperties
	IsHeading() bool
	HeadingLevel() int
	Alignment() string
	SetAlignment(align string)
	SpacingBefore() int64
	SetSpacingBefore(twips int64)
	SpacingAfter() int64
	SetSpacingAfter(twips int64)
	KeepWithNext() bool
	SetKeepWithNext(v bool)
	ListLevel() int
	SetListLevel(level int) error
	ListNumberingID() int
	SetListNumberingID(numID int)
	SetList(numID, level int) error
	AddContentControl(tag, alias, text string) *ContentControl
	AddHyperlink(url, text string) (*Hyperlink, error)
	AddHyperlinkWithTooltip(url, text, tooltip string) (*Hyperlink, error)
	AddBookmarkLink(anchor, text string) (*Hyperlink, error)
	AddBookmark(name string, startRun, endRun int) error
	AddField(instruction, display string) (*Field, error)
	Hyperlinks() []*Hyperlink
	ContentControls() []*ContentControl
}

// Run represents a text run.
type Run interface {
	Text() string
	SetText(text string)
	Bold() bool
	SetBold(v bool)
	Italic() bool
	SetItalic(v bool)
	Underline() bool
	SetUnderline(v bool)
	UnderlineStyle() string
	SetUnderlineStyle(style string)
	Strike() bool
	SetStrike(v bool)
	FontSize() float64
	SetFontSize(points float64)
	FontName() string
	SetFontName(name string)
	Color() string
	SetColor(hex string)
	Highlight() string
	SetHighlight(color string)
	Superscript() bool
	SetSuperscript(v bool)
	Subscript() bool
	SetSubscript(v bool)
	Properties() RunProperties
	AddBreak()
	AddPageBreak()
	AddTab()
}

// Table represents a table.
type Table interface {
	BodyElement
	Rows() []Row
	Row(index int) Row
	AddRow() Row
	InsertRow(index int) Row
	DeleteRow(index int) error
	Cell(row, col int) Cell
	ColumnCount() int
	RowCount() int
	FirstRowText() []string
	Purpose() string
	Style() string
	SetStyle(styleID string)
}

// Row represents a table row.
type Row interface {
	Cells() []Cell
	Cell(index int) Cell
	AddCell() Cell
	Index() int
	IsHeader() bool
	SetHeader(v bool)
}

// Cell represents a table cell.
type Cell interface {
	Text() string
	SetText(text string)
	Paragraphs() []Paragraph
	AddParagraph() Paragraph
	VerticalMerge() VerticalMerge
	SetVerticalMerge(v VerticalMerge)
	GridSpan() int
	SetGridSpan(span int)
	Shading() string
	SetShading(fill string)
	Index() int
}
