package packaging

import (
	"bytes"
	"io"
)

// Part represents a part within an OPC package.
type Part struct {
	uri         string
	contentType string
	content     []byte
	pkg         *Package
	modified    bool
}

// URI returns the part's URI within the package.
func (p *Part) URI() string {
	return p.uri
}

// ContentType returns the part's content type.
func (p *Part) ContentType() string {
	return p.contentType
}

// Content returns the part's content as a byte slice.
func (p *Part) Content() ([]byte, error) {
	return p.content, nil
}

// SetContent sets the part's content.
func (p *Part) SetContent(content []byte) error {
	p.content = content
	p.modified = true
	return nil
}

// Stream returns the part's content as an io.ReadCloser.
func (p *Part) Stream() (io.ReadCloser, error) {
	return io.NopCloser(bytes.NewReader(p.content)), nil
}

// Size returns the size of the part's content.
func (p *Part) Size() int {
	return len(p.content)
}

// IsModified returns true if the part has been modified.
func (p *Part) IsModified() bool {
	return p.modified
}

// newPart creates a new part.
func newPart(uri, contentType string, content []byte, pkg *Package) *Part {
	return &Part{
		uri:         uri,
		contentType: contentType,
		content:     content,
		pkg:         pkg,
		modified:    false,
	}
}
