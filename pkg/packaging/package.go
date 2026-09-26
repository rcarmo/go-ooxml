package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"sort"
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
	if filePath == "" {
		return nil, utils.ErrPathNotSet
	}
	cleanPath := filepath.Clean(filePath)
	f, err := os.Open(cleanPath)
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
	pkg.path = cleanPath
	return pkg, nil
}

// OpenReader opens an OPC package from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (*Package, error) {
	return OpenReaderWithLimits(r, size, Limits{})
}

func openReader(r io.ReaderAt, size int64, limits Limits) (*Package, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, err
	}

	pkg := &Package{
		parts:         make(map[string]*Part),
		relationships: make(map[string]*Relationships),
	}

	if err := validateBudgets(zr.File, limits); err != nil {
		return nil, err
	}
	if err := validateZIPStructure(r, size, zr.File); err != nil {
		return nil, err
	}

	// Validate names before inflating or inserting members into a map. Never
	// silently let the last duplicate win.
	seen := make(map[string]string, len(zr.File))
	for _, f := range zr.File {
		if err := validateMemberName(f.Name, f.FileInfo().IsDir()); err != nil {
			return nil, err
		}
		key := strings.ToLower(strings.TrimSuffix(f.Name, "/"))
		if prior, exists := seen[key]; exists {
			return nil, invalidPart("open", f.Name, "duplicate or case-colliding member with "+prior)
		}
		seen[key] = f.Name
		if f.Flags&1 != 0 {
			return nil, invalidPart("open", f.Name, "encrypted member")
		}
		if f.Method != zip.Store && f.Method != zip.Deflate {
			return nil, invalidPart("open", f.Name, "unsupported compression")
		}
	}

	// Read all files from ZIP. Directory entries are not OPC parts.
	for _, f := range zr.File {
		if f.FileInfo().IsDir() {
			continue
		}
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
		return utils.ErrPathNotSet
	}
	return p.SaveAs(p.path)
}

// SaveAs saves the package to a new path.
func (p *Package) SaveAs(filePath string) error {
	if p.closed {
		return utils.ErrDocumentClosed
	}

	if filePath == "" {
		return utils.ErrPathNotSet
	}
	cleanPath := filepath.Clean(filePath)
	// A sibling temporary file keeps a serialization failure from truncating
	// an existing destination. New files are private; replacements retain mode.
	mode := os.FileMode(0600)
	if info, err := os.Stat(cleanPath); err == nil {
		if !info.Mode().IsRegular() {
			return fmt.Errorf("destination is not a regular file: %s", cleanPath)
		}
		mode = info.Mode().Perm()
	} else if !os.IsNotExist(err) {
		return err
	}
	f, err := os.CreateTemp(filepath.Dir(cleanPath), ".ooxml-*")
	if err != nil {
		return err
	}
	temporary := f.Name()
	defer func() { _ = f.Close(); _ = os.Remove(temporary) }()
	if err := f.Chmod(mode); err != nil {
		return err
	}
	if err := p.WriteTo(f); err != nil {
		return err
	}
	if err := f.Sync(); err != nil {
		return err
	}
	if err := f.Close(); err != nil {
		return err
	}
	if err := os.Rename(temporary, cleanPath); err != nil {
		return err
	}

	p.path = cleanPath
	p.modified = false
	return nil
}

// WriteTo writes the package to an io.Writer.
func (p *Package) WriteTo(w io.Writer) error {
	if p.closed {
		return utils.ErrDocumentClosed
	}
	if err := p.validateOutputNames(); err != nil {
		return err
	}
	zw := zip.NewWriter(w)

	// Write [Content_Types].xml first
	ctData, err := xml.Marshal(p.contentTypes)
	if err != nil {
		return err
	}
	ctData = append([]byte(utils.XMLHeader), ctData...)
	if err := writeZipFile(zw, ContentTypesPath, ctData); err != nil {
		return err
	}

	// Write relationship files and parts in stable order.
	sources := make([]string, 0, len(p.relationships))
	for source := range p.relationships {
		sources = append(sources, source)
	}
	sort.Strings(sources)
	for _, sourceURI := range sources {
		rels := p.relationships[sourceURI]
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

	// Write all parts.
	uris := make([]string, 0, len(p.parts))
	for uri := range p.parts {
		uris = append(uris, uri)
	}
	sort.Strings(uris)
	for _, uri := range uris {
		part := p.parts[uri]
		// Skip [Content_Types].xml and .rels files (already written)
		if uri == ContentTypesPath || strings.HasSuffix(uri, ".rels") {
			continue
		}
		if err := writeZipFile(zw, uri, part.content); err != nil {
			return err
		}
	}

	// Close emits the central directory and flushes buffered bytes. Its error
	// is part of the write result, not a best-effort cleanup detail.
	return zw.Close()
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
	return p.AddRelationshipWithTargetMode(sourceURI, targetURI, relType, TargetModeInternal)
}

// AddRelationshipWithTargetMode adds a relationship with the specified target mode.
func (p *Package) AddRelationshipWithTargetMode(sourceURI, targetURI, relType string, targetMode TargetMode) *Relationship {
	sourceURI = normalizePath(sourceURI)
	rels := p.GetRelationships(sourceURI)
	id := rels.NextID()
	rel := rels.AddWithID(id, relType, targetURI, targetMode)
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
		return utils.ErrMissingContentTypes
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
		return nil, invalidPart("open", f.Name, err.Error())
	}
	defer rc.Close()
	// Read at most the declared length plus one byte, even if a malicious
	// stream expands beyond its advertised size. ReadAll must reach checksum
	// verification before any bytes enter the package model.
	content, err := io.ReadAll(io.LimitReader(rc, int64(f.UncompressedSize64)+1))
	if err != nil {
		return nil, invalidPart("open", f.Name, err.Error())
	}
	if uint64(len(content)) != f.UncompressedSize64 {
		return nil, invalidPart("open", f.Name, "declared and actual sizes differ")
	}
	return content, nil
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
