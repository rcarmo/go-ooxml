// Package document provides hyperlink functionality.
package document

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Hyperlink represents a hyperlink in a paragraph.
type Hyperlink struct {
	doc *documentImpl
	h   *wml.Hyperlink
}

// URL returns the hyperlink target URL if available.
func (h *Hyperlink) URL() string {
	if h.doc == nil || h.doc.pkg == nil || h.h == nil || h.h.ID == "" {
		return ""
	}
	rels := h.doc.pkg.GetRelationships(packaging.WordDocumentPath)
	rel := rels.ByID(h.h.ID)
	if rel == nil || rel.TargetMode != packaging.TargetModeExternal {
		return ""
	}
	return rel.Target
}

// Anchor returns the hyperlink anchor (bookmark name).
func (h *Hyperlink) Anchor() string {
	return h.h.Anchor
}

// Tooltip returns the hyperlink tooltip text.
func (h *Hyperlink) Tooltip() string {
	return h.h.Tooltip
}

// Text returns the hyperlink text.
func (h *Hyperlink) Text() string {
	return textFromInlineContent(h.h.Content)
}

// AddHyperlink adds a hyperlink with display text to the paragraph.
func (p *paragraphImpl) AddHyperlink(url, text string) (*Hyperlink, error) {
	if url == "" {
		return nil, utils.NewValidationError("url", "cannot be empty", url)
	}

	rel := p.doc.pkg.AddRelationshipWithTargetMode(packaging.WordDocumentPath, url, packaging.RelTypeHyperlink, packaging.TargetModeExternal)
	link := &wml.Hyperlink{
		ID: rel.ID,
		Content: []interface{}{
			&wml.R{Content: []interface{}{wml.NewT(text)}},
		},
	}
	p.p.Content = append(p.p.Content, link)
	return &Hyperlink{doc: p.doc, h: link}, nil
}

// AddHyperlinkWithTooltip adds a hyperlink with tooltip text.
func (p *paragraphImpl) AddHyperlinkWithTooltip(url, text, tooltip string) (*Hyperlink, error) {
	link, err := p.AddHyperlink(url, text)
	if err != nil {
		return nil, err
	}
	link.h.Tooltip = tooltip
	return link, nil
}

// AddBookmarkLink adds a hyperlink pointing to a bookmark anchor.
func (p *paragraphImpl) AddBookmarkLink(anchor, text string) (*Hyperlink, error) {
	if anchor == "" {
		return nil, utils.NewValidationError("anchor", "cannot be empty", anchor)
	}
	link := &wml.Hyperlink{
		Anchor:  anchor,
		Content: []interface{}{
			&wml.R{Content: []interface{}{wml.NewT(text)}},
		},
	}
	p.p.Content = append(p.p.Content, link)
	return &Hyperlink{doc: p.doc, h: link}, nil
}
