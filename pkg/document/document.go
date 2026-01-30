// Package document provides a high-level API for working with Word documents.
package document

import (
	"io"
	"time"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Document represents a Word document.
type Document struct {
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
	headers map[string]*Header
	footers map[string]*Footer
}

// New creates a new empty Word document.
func New() (*Document, error) {
	doc := &Document{
		pkg: packaging.New(),
		document: &wml.Document{
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
		headers:  make(map[string]*Header),
		footers:  make(map[string]*Footer),
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
func Open(path string) (*Document, error) {
	pkg, err := packaging.Open(path)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

// OpenReader opens a Word document from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (*Document, error) {
	pkg, err := packaging.OpenReader(r, size)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

func openFromPackage(pkg *packaging.Package) (*Document, error) {
	doc := &Document{
		pkg:     pkg,
		headers: make(map[string]*Header),
		footers: make(map[string]*Footer),
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

	return doc, nil
}

// Save saves the document to its original path.
func (d *Document) Save() error {
	if err := d.updatePackage(); err != nil {
		return err
	}
	return d.pkg.Save()
}

// SaveAs saves the document to a new path.
func (d *Document) SaveAs(path string) error {
	if err := d.updatePackage(); err != nil {
		return err
	}
	return d.pkg.SaveAs(path)
}

// Close closes the document.
func (d *Document) Close() error {
	return d.pkg.Close()
}

// Body returns the document body.
func (d *Document) Body() *Body {
	return &Body{doc: d}
}

// Paragraphs returns all paragraphs in the document.
func (d *Document) Paragraphs() []*Paragraph {
	return d.Body().Paragraphs()
}

// Tables returns all tables in the document.
func (d *Document) Tables() []*Table {
	return d.Body().Tables()
}

// AddParagraph adds a new paragraph to the document body.
func (d *Document) AddParagraph() *Paragraph {
	return d.Body().AddParagraph()
}

// AddTable adds a new table to the document body.
func (d *Document) AddTable(rows, cols int) *Table {
	return d.Body().AddTable(rows, cols)
}

// TrackChangesEnabled returns whether track changes is enabled.
func (d *Document) TrackChangesEnabled() bool {
	return d.trackChanges
}

// EnableTrackChanges enables track changes.
func (d *Document) EnableTrackChanges(author string) {
	d.trackChanges = true
	d.trackAuthor = author
	if d.settings.TrackRevisions == nil {
		d.settings.TrackRevisions = wml.NewOnOffEnabled()
	}
}

// DisableTrackChanges disables track changes.
func (d *Document) DisableTrackChanges() {
	d.trackChanges = false
	d.settings.TrackRevisions = nil
}

// TrackAuthor returns the author name for tracked changes.
func (d *Document) TrackAuthor() string {
	return d.trackAuthor
}

// SetTrackAuthor sets the author name for tracked changes.
func (d *Document) SetTrackAuthor(author string) {
	d.trackAuthor = author
}

// nextRevID returns the next revision ID.
func (d *Document) nextRevID() int {
	d.nextRevisionID++
	return d.nextRevisionID
}

// nextCommID returns the next comment ID.
func (d *Document) nextCommID() int {
	d.nextCommentID++
	return d.nextCommentID
}

// currentDateTime returns the current date/time in OOXML format.
func currentDateTime() string {
	return time.Now().UTC().Format("2006-01-02T15:04:05Z")
}
