// Package document provides a high-level API for working with Word documents.
package document

import (
	"encoding/xml"
	"io"
	"strconv"
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// documentImpl represents a Word document.
type documentImpl struct {
	pkg      *packaging.Package
	document *wml.Document
	styles   *wml.Styles
	settings *wml.Settings
	comments *wml.Comments
	numbering *wml.Numbering

	// Tracking
	trackChanges     bool
	trackAuthor      string
	nextRevisionID   int
	nextCommentID    int
	nextAbstractNumID int
	nextNumID         int
	nextBookmarkID    int

	// Headers and footers (keyed by relID)
	headers map[string]*headerImpl
	footers map[string]*footerImpl
}

// New creates a new empty Word document.
func New() (Document, error) {
	doc := &documentImpl{
		pkg: packaging.New(),
		document: &wml.Document{
			XMLName: xml.Name{Space: wml.NS, Local: "document"},
			MCIgnorable: "",
			Body: &wml.Body{
				SectPr: &wml.SectPr{
					PgSz: &wml.PgSz{W: 12240, H: 15840}, // Letter size
					PgMar: &wml.PgMar{
						Top:    1440, Right: 1440, Bottom: 1440, Left: 1440,
						Header: 720, Footer: 720, Gutter: 0,
					},
				},
			},
		},
		styles:   &wml.Styles{},
		settings: &wml.Settings{},
		headers:  make(map[string]*headerImpl),
		footers:  make(map[string]*footerImpl),
		nextAbstractNumID: 1,
		nextNumID:         1,
		nextBookmarkID:    1,
	}

	// Initialize package structure
	if err := doc.initPackage(); err != nil {
		return nil, err
	}

	return doc, nil
}

// Open opens an existing Word document.
func Open(path string) (Document, error) {
	pkg, err := packaging.Open(path)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

// OpenReader opens a Word document from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (Document, error) {
	pkg, err := packaging.OpenReader(r, size)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

func openFromPackage(pkg *packaging.Package) (*documentImpl, error) {
	doc := &documentImpl{
		pkg:     pkg,
		headers: make(map[string]*headerImpl),
		footers: make(map[string]*footerImpl),
	}

	// Parse document.xml
	if err := doc.parseDocument(); err != nil {
		return nil, err
	}

	// Parse styles.xml (optional)
	_ = doc.parseStyles()

	// Parse settings.xml (optional)
	_ = doc.parseSettings()

	// Parse comments.xml (optional)
	_ = doc.parseComments()
	// Parse numbering.xml (optional)
	_ = doc.parseNumbering()
	doc.parseBookmarks()
	_ = doc.parseHeaders()
	_ = doc.parseFooters()

	return doc, nil
}

// Save saves the document to its original path.
func (d *documentImpl) Save() error {
	if err := d.updatePackage(); err != nil {
		return err
	}
	return d.pkg.Save()
}

// SaveAs saves the document to a new path.
func (d *documentImpl) SaveAs(path string) error {
	if err := d.updatePackage(); err != nil {
		return err
	}
	return d.pkg.SaveAs(path)
}

// Close closes the document.
func (d *documentImpl) Close() error {
	return d.pkg.Close()
}

// Body returns the document body.
func (d *documentImpl) Body() Body {
	return &bodyImpl{doc: d}
}

// Paragraphs returns all paragraphs in the document.
func (d *documentImpl) Paragraphs() []Paragraph {
	return d.Body().Paragraphs()
}

// Tables returns all tables in the document.
func (d *documentImpl) Tables() []Table {
	return d.Body().Tables()
}

// Sections returns the document sections.
func (d *documentImpl) Sections() []Section {
	if d == nil || d.document == nil {
		return nil
	}
	if d.document.Body == nil {
		d.document.Body = &wml.Body{}
	}
	if d.document.Body.SectPr == nil {
		d.document.Body.SectPr = &wml.SectPr{}
	}
	return []Section{&sectionImpl{doc: d, sectPr: d.document.Body.SectPr}}
}

// XML returns the underlying WML document for advanced access.
func (d *documentImpl) XML() *wml.Document {
	return d.document
}

// CoreProperties returns the document core properties.
func (d *documentImpl) CoreProperties() (*common.CoreProperties, error) {
	return d.pkg.CoreProperties()
}

// SetCoreProperties sets the document core properties.
func (d *documentImpl) SetCoreProperties(props *common.CoreProperties) error {
	return d.pkg.SetCoreProperties(props)
}

// Properties returns the document properties (core properties).
func (d *documentImpl) Properties() DocumentProperties {
	props, _ := d.pkg.CoreProperties()
	if props == nil {
		return DocumentProperties{}
	}
	return *props
}

// AddParagraph adds a new paragraph to the document body.
func (d *documentImpl) AddParagraph() Paragraph {
	return d.Body().AddParagraph()
}

// AddTable adds a new table to the document body.
func (d *documentImpl) AddTable(rows, cols int) Table {
	return d.Body().AddTable(rows, cols)
}

// TrackChangesEnabled reports whether track changes is enabled.
func (d *documentImpl) TrackChangesEnabled() bool {
	return d.trackChanges
}

// EnableTrackChanges enables track changes with an author.
func (d *documentImpl) EnableTrackChanges(author string) {
	d.trackChanges = true
	d.trackAuthor = author
	if d.settings.TrackRevisions == nil {
		d.settings.TrackRevisions = wml.NewOnOffEnabled()
	}
}

// DisableTrackChanges disables track changes.
func (d *documentImpl) DisableTrackChanges() {
	d.trackChanges = false
	d.settings.TrackRevisions = nil
}

// TrackAuthor returns the author name for tracked changes.
func (d *documentImpl) TrackAuthor() string {
	return d.trackAuthor
}

// SetTrackAuthor sets the author name for tracked changes.
func (d *documentImpl) SetTrackAuthor(author string) {
	d.trackAuthor = author
}

// TrackChanges returns the track changes manager.
func (d *documentImpl) TrackChanges() TrackChanges {
	return &TrackChangesManager{doc: d}
}

// Insertions returns tracked insertions.
func (d *documentImpl) Insertions() []Revision {
	return d.revisionsByType(RevisionInsert)
}

// Deletions returns tracked deletions.
func (d *documentImpl) Deletions() []Revision {
	return d.revisionsByType(RevisionDelete)
}

// StyleByID returns a style by ID.
func (d *documentImpl) StyleByID(id string) Style {
	if d == nil {
		return nil
	}
	return d.Styles().ByID(id)
}

// StyleByName returns a style by name.
func (d *documentImpl) StyleByName(name string) Style {
	if d == nil {
		return nil
	}
	return d.Styles().ByName(name)
}


// AcceptAllRevisions accepts all tracked revisions.
func (d *documentImpl) AcceptAllRevisions() {
	for _, rev := range d.AllRevisions() {
		_ = rev.Accept()
	}
}

// RejectAllRevisions rejects all tracked revisions.
func (d *documentImpl) RejectAllRevisions() {
	for _, rev := range d.AllRevisions() {
		_ = rev.Reject()
	}
}

// CommentByID returns a comment by its ID.
func (d *documentImpl) CommentByID(id string) Comment {
	if d.comments == nil {
		return nil
	}
	for _, c := range d.comments.Comment {
		if strconv.Itoa(c.ID) == id {
			return &commentImpl{doc: d, comment: c}
		}
	}
	return nil
}

// DeleteComment removes a comment by ID.
func (d *documentImpl) DeleteComment(id string) error {
	if d.comments == nil {
		return utils.ErrCommentNotFound
	}
	for i, c := range d.comments.Comment {
		if strconv.Itoa(c.ID) == id {
			d.comments.Comment = append(d.comments.Comment[:i], d.comments.Comment[i+1:]...)
			commentID, _ := strconv.Atoi(id)
			d.removeCommentRanges(commentID)
			return nil
		}
	}
	return utils.ErrCommentNotFound
}

// nextRevID returns the next revision ID.
func (d *documentImpl) nextRevID() int {
	d.nextRevisionID++
	return d.nextRevisionID
}

// nextCommID returns the next comment ID.
func (d *documentImpl) nextCommID() int {
	d.nextCommentID++
	return d.nextCommentID
}

// currentDateTime returns the current date/time in OOXML format.
func currentDateTime() string {
	return time.Now().UTC().Format("2006-01-02T15:04:05Z")
}
