package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"os"
	"path"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Package represents an OPC package (ZIP archive with relationships).
type Package struct {
	path          string
	contentTypes  *ContentTypes
	parts         map[string]*Part
	relationships map[string]*Relationships // key is source part URI ("" for package-level)
	closed        bool
	modified      bool
}

// Open opens an existing OPC package from a file path.
func Open(filePath string) (*Package, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	stat, err := f.Stat()
	if err != nil {
		return nil, err
	}

	pkg, err := OpenReader(f, stat.Size())
	if err != nil {
		return nil, err
	}
	pkg.path = filePath
	return pkg, nil
}

// OpenReader opens an OPC package from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (*Package, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, err
	}

	pkg := &Package{
		parts:         make(map[string]*Part),
		relationships: make(map[string]*Relationships),
	}

	// Read all files from ZIP
	for _, f := range zr.File {
		content, err := readZipFile(f)
		if err != nil {
			return nil, err
		}

		// Normalize path (remove leading /)
		uri := strings.TrimPrefix(f.Name, "/")
		pkg.parts[uri] = newPart(uri, "", content, pkg)
	}

	// Parse [Content_Types].xml
	if err := pkg.parseContentTypes(); err != nil {
		return nil, err
	}

	// Set content types on parts
	for uri, part := range pkg.parts {
		part.contentType = pkg.contentTypes.GetContentType(uri)
	}

	// Parse relationships
	if err := pkg.parseRelationships(); err != nil {
		return nil, err
	}

	return pkg, nil
}

// OpenBytes opens an OPC package from a byte slice.
func OpenBytes(data []byte) (*Package, error) {
	return OpenReader(bytes.NewReader(data), int64(len(data)))
}

// New creates a new empty OPC package.
func New() *Package {
	return &Package{
		contentTypes:  NewContentTypes(),
		parts:         make(map[string]*Part),
		relationships: make(map[string]*Relationships),
		modified:      true,
	}
}

// Save saves the package to its original path.
func (p *Package) Save() error {
	if p.path == "" {
		return errors.New("no file path set; use SaveAs")
	}
	return p.SaveAs(p.path)
}

// SaveAs saves the package to a new path.
func (p *Package) SaveAs(filePath string) error {
	if p.closed {
		return utils.ErrDocumentClosed
	}

	f, err := os.Create(filePath)
	if err != nil {
		return err
	}
	defer f.Close()

	if err := p.WriteTo(f); err != nil {
		return err
	}

	p.path = filePath
	p.modified = false
	return nil
}

// WriteTo writes the package to an io.Writer.
func (p *Package) WriteTo(w io.Writer) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write [Content_Types].xml first
	ctData, err := xml.Marshal(p.contentTypes)
	if err != nil {
		return err
	}
	ctData = append([]byte(utils.XMLHeader), ctData...)
	if err := writeZipFile(zw, ContentTypesPath, ctData); err != nil {
		return err
	}

	// Write relationships files
	for sourceURI, rels := range p.relationships {
		if len(rels.Relationships) == 0 {
			continue
		}
		var relsPath string
		if sourceURI == "" || sourceURI == "." {
			relsPath = PackageRelsPath
		} else {
			relsPath = RelationshipsPathForPart(sourceURI)
		}
		relsData, err := xml.Marshal(rels)
		if err != nil {
			return err
		}
		relsData = append([]byte(utils.XMLHeader), relsData...)
		if err := writeZipFile(zw, relsPath, relsData); err != nil {
			return err
		}
	}

	// Write all parts
	for uri, part := range p.parts {
		// Skip [Content_Types].xml and .rels files (already written)
		if uri == ContentTypesPath || strings.HasSuffix(uri, ".rels") {
			continue
		}
		if err := writeZipFile(zw, uri, part.content); err != nil {
			return err
		}
	}

	return nil
}

// Close closes the package.
func (p *Package) Close() error {
	p.closed = true
	p.parts = nil
	p.relationships = nil
	return nil
}

// GetPart returns a part by URI.
func (p *Package) GetPart(uri string) (*Part, error) {
	if p.closed {
		return nil, utils.ErrDocumentClosed
	}
	uri = normalizePath(uri)
	part, ok := p.parts[uri]
	if !ok {
		return nil, utils.ErrPartNotFound
	}
	return part, nil
}

// AddPart adds a new part to the package.
func (p *Package) AddPart(uri, contentType string, content []byte) (*Part, error) {
	if p.closed {
		return nil, utils.ErrDocumentClosed
	}
	uri = normalizePath(uri)
	part := newPart(uri, contentType, content, p)
	part.modified = true
	p.parts[uri] = part
	p.contentTypes.EnsureContentType(uri, contentType)
	p.modified = true
	return part, nil
}

// DeletePart removes a part from the package.
func (p *Package) DeletePart(uri string) error {
	if p.closed {
		return utils.ErrDocumentClosed
	}
	uri = normalizePath(uri)
	if _, ok := p.parts[uri]; !ok {
		return utils.ErrPartNotFound
	}
	delete(p.parts, uri)
	p.contentTypes.RemoveOverride(uri)
	p.modified = true
	return nil
}

// Parts returns all parts in the package.
func (p *Package) Parts() []*Part {
	result := make([]*Part, 0, len(p.parts))
	for _, part := range p.parts {
		result = append(result, part)
	}
	return result
}

// PartExists checks if a part exists.
func (p *Package) PartExists(uri string) bool {
	uri = normalizePath(uri)
	_, ok := p.parts[uri]
	return ok
}

// GetRelationships returns relationships for a source part URI.
// Use empty string for package-level relationships.
func (p *Package) GetRelationships(sourceURI string) *Relationships {
	sourceURI = normalizePath(sourceURI)
	rels, ok := p.relationships[sourceURI]
	if !ok {
		rels = NewRelationships()
		p.relationships[sourceURI] = rels
	}
	return rels
}

// AddRelationship adds a relationship from a source part.
func (p *Package) AddRelationship(sourceURI, targetURI, relType string) *Relationship {
	sourceURI = normalizePath(sourceURI)
	rels := p.GetRelationships(sourceURI)
	rel := rels.Add(relType, targetURI, TargetModeInternal)
	p.modified = true
	return rel
}

// GetRelationshipsByType returns relationships of a specific type.
func (p *Package) GetRelationshipsByType(sourceURI, relType string) []*Relationship {
	return p.GetRelationships(sourceURI).ByType(relType)
}

// GetContentType returns the content type for a part URI.
func (p *Package) GetContentType(uri string) string {
	return p.contentTypes.GetContentType(uri)
}

// ContentTypes returns the package's content types.
func (p *Package) ContentTypes() *ContentTypes {
	return p.contentTypes
}

// Path returns the package's file path.
func (p *Package) Path() string {
	return p.path
}

// IsModified returns true if the package has been modified.
func (p *Package) IsModified() bool {
	return p.modified
}

// parseContentTypes reads and parses [Content_Types].xml.
func (p *Package) parseContentTypes() error {
	part, ok := p.parts[ContentTypesPath]
	if !ok {
		return errors.New("missing [Content_Types].xml")
	}

	p.contentTypes = &ContentTypes{}
	if err := utils.UnmarshalXML(part.content, p.contentTypes); err != nil {
		return err
	}
	return nil
}

// parseRelationships reads and parses all .rels files.
func (p *Package) parseRelationships() error {
	for uri, part := range p.parts {
		if !strings.HasSuffix(uri, ".rels") {
			continue
		}

		rels := &Relationships{}
		if err := utils.UnmarshalXML(part.content, rels); err != nil {
			return err
		}

		// Determine source part URI
		sourceURI := ""
		if uri != PackageRelsPath {
			// Extract source part from .rels path
			// e.g., "word/_rels/document.xml.rels" -> "word/document.xml"
			dir := path.Dir(path.Dir(uri))
			base := strings.TrimSuffix(path.Base(uri), ".rels")
			if dir == "." {
				sourceURI = base
			} else {
				sourceURI = dir + "/" + base
			}
		}
		sourceURI = normalizePath(sourceURI)

		p.relationships[sourceURI] = rels
	}
	return nil
}

// Helper functions

func readZipFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

func writeZipFile(zw *zip.Writer, name string, data []byte) error {
	w, err := zw.Create(name)
	if err != nil {
		return err
	}
	_, err = w.Write(data)
	return err
}

func normalizePath(p string) string {
	p = strings.TrimPrefix(p, "/")
	return path.Clean(p)
}
