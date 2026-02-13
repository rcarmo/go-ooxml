// Package document provides header and footer functionality.
package document

import (
	"fmt"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// HeaderFooterType represents the type of header/footer.
type HeaderFooterType string

const (
	// HeaderFooterDefault identifies the default header/footer.
	HeaderFooterDefault HeaderFooterType = "default"
	// HeaderFooterFirst identifies the first-page header/footer.
	HeaderFooterFirst HeaderFooterType = "first"
	// HeaderFooterEven identifies the even-page header/footer.
	HeaderFooterEven HeaderFooterType = "even"
)

// =============================================================================
// Header methods
// =============================================================================

// Type returns the header type.
func (h *headerImpl) Type() HeaderFooterType {
	return h.hfType
}

// Paragraphs returns all paragraphs in the header.
func (h *headerImpl) Paragraphs() []Paragraph {
	if h.header == nil {
		return nil
	}

	var result []Paragraph
	for i, p := range h.header.Content {
		if para, ok := p.(*wml.P); ok {
			result = append(result, &paragraphImpl{doc: h.doc, p: para, index: i})
		}
	}
	return result
}

// AddParagraph adds a new paragraph to the header.
func (h *headerImpl) AddParagraph() Paragraph {
	if h.header == nil {
		h.header = &wml.Header{}
	}

	p := &wml.P{}
	h.header.Content = append(h.header.Content, p)
	return &paragraphImpl{doc: h.doc, p: p, index: len(h.header.Content) - 1}
}

// Text returns the combined text of all paragraphs.
func (h *headerImpl) Text() string {
	var text string
	for _, para := range h.Paragraphs() {
		if text != "" {
			text += "\n"
		}
		text += para.Text()
	}
	return text
}

// SetText sets the header text, replacing all content.
func (h *headerImpl) SetText(text string) {
	h.header.Content = []interface{}{&wml.P{}}
	if len(h.Paragraphs()) > 0 {
		h.Paragraphs()[0].SetText(text)
	}
}

// =============================================================================
// Footer methods
// =============================================================================

// Type returns the footer type.
func (f *footerImpl) Type() HeaderFooterType {
	return f.hfType
}

// Paragraphs returns all paragraphs in the footer.
func (f *footerImpl) Paragraphs() []Paragraph {
	if f.footer == nil {
		return nil
	}

	var result []Paragraph
	for i, p := range f.footer.Content {
		if para, ok := p.(*wml.P); ok {
			result = append(result, &paragraphImpl{doc: f.doc, p: para, index: i})
		}
	}
	return result
}

// AddParagraph adds a new paragraph to the footer.
func (f *footerImpl) AddParagraph() Paragraph {
	if f.footer == nil {
		f.footer = &wml.Footer{}
	}

	p := &wml.P{}
	f.footer.Content = append(f.footer.Content, p)
	return &paragraphImpl{doc: f.doc, p: p, index: len(f.footer.Content) - 1}
}

// Text returns the combined text of all paragraphs.
func (f *footerImpl) Text() string {
	var text string
	for _, para := range f.Paragraphs() {
		if text != "" {
			text += "\n"
		}
		text += para.Text()
	}
	return text
}

// SetText sets the footer text, replacing all content.
func (f *footerImpl) SetText(text string) {
	f.footer.Content = []interface{}{&wml.P{}}
	if len(f.Paragraphs()) > 0 {
		f.Paragraphs()[0].SetText(text)
	}
}

// =============================================================================
// Document header/footer methods
// =============================================================================

// Headers returns all headers in the document.
func (d *documentImpl) Headers() []Header {
	var result []Header

	sectPr := d.document.Body.SectPr
	if sectPr == nil {
		return nil
	}

	for _, ref := range sectPr.HeaderRefs {
		h := d.headerByRelID(ref.ID)
		if h != nil {
			h.hfType = HeaderFooterType(ref.Type)
			result = append(result, h)
		}
	}

	return result
}

// Footers returns all footers in the document.
func (d *documentImpl) Footers() []Footer {
	var result []Footer

	sectPr := d.document.Body.SectPr
	if sectPr == nil {
		return nil
	}

	for _, ref := range sectPr.FooterRefs {
		f := d.footerByRelID(ref.ID)
		if f != nil {
			f.hfType = HeaderFooterType(ref.Type)
			result = append(result, f)
		}
	}

	return result
}

// Header returns the header of the specified type.
func (d *documentImpl) Header(hfType HeaderFooterType) Header {
	for _, header := range d.Headers() {
		if h, ok := header.(*headerImpl); ok && h.hfType == hfType {
			return h
		}
	}
	return nil
}

// Footer returns the footer of the specified type.
func (d *documentImpl) Footer(hfType HeaderFooterType) Footer {
	for _, footer := range d.Footers() {
		if f, ok := footer.(*footerImpl); ok && f.hfType == hfType {
			return f
		}
	}
	return nil
}

// AddHeader adds a header of the specified type.
func (d *documentImpl) AddHeader(hfType HeaderFooterType) Header {
	// Generate unique filename and relationship ID
	num := maxPartCounter(d.pkg, "word/header", ".xml")
	filename := fmt.Sprintf("header%d.xml", num)
	rels := d.pkg.GetRelationships(packaging.WordDocumentPath)
	relID := rels.NextID()
	
	header := &wml.Header{
		Content: []interface{}{&wml.P{}},
	}
	
	// Add to section properties
	if d.document.Body.SectPr == nil {
		d.document.Body.SectPr = &wml.SectPr{}
	}
	
	d.document.Body.SectPr.HeaderRefs = append(d.document.Body.SectPr.HeaderRefs, wml.HeaderRef{
		Type: string(hfType),
		ID:   relID,
	})
	
	// Store and add relationship
	h := &headerImpl{
		doc:    d,
		header: header,
		relID:  relID,
		hfType: hfType,
	}
	
	d.headers[relID] = h
	rels.AddWithID(relID, packaging.RelTypeHeader, filename, packaging.TargetModeInternal)
	
	return h
}

// AddFooter adds a footer of the specified type.
func (d *documentImpl) AddFooter(hfType HeaderFooterType) Footer {
	// Generate unique filename and relationship ID
	num := maxPartCounter(d.pkg, "word/footer", ".xml")
	filename := fmt.Sprintf("footer%d.xml", num)
	rels := d.pkg.GetRelationships(packaging.WordDocumentPath)
	relID := rels.NextID()
	
	footer := &wml.Footer{
		Content: []interface{}{&wml.P{}},
	}
	
	// Add to section properties
	if d.document.Body.SectPr == nil {
		d.document.Body.SectPr = &wml.SectPr{}
	}
	
	d.document.Body.SectPr.FooterRefs = append(d.document.Body.SectPr.FooterRefs, wml.FooterRef{
		Type: string(hfType),
		ID:   relID,
	})
	
	// Store and add relationship
	f := &footerImpl{
		doc:    d,
		footer: footer,
		relID:  relID,
		hfType: hfType,
	}
	
	d.footers[relID] = f
	rels.AddWithID(relID, packaging.RelTypeFooter, filename, packaging.TargetModeInternal)
	
	return f
}

func (d *documentImpl) headerByRelID(relID string) *headerImpl {
	if h, ok := d.headers[relID]; ok {
		return h
	}
	// Return empty header for unknown relIDs (e.g., from loaded documents)
	return &headerImpl{doc: d, relID: relID}
}

func (d *documentImpl) footerByRelID(relID string) *footerImpl {
	if f, ok := d.footers[relID]; ok {
		return f
	}
	// Return empty footer for unknown relIDs (e.g., from loaded documents)
	return &footerImpl{doc: d, relID: relID}
}

// =============================================================================
// Header/Footer marshaling
// =============================================================================

func (d *documentImpl) saveHeaders() error {
	if d.document == nil || d.document.Body == nil || d.document.Body.SectPr == nil {
		return nil
	}

	rels := d.pkg.GetRelationships(packaging.WordDocumentPath)
	i := 0
	for _, ref := range d.document.Body.SectPr.HeaderRefs {
		h := d.headers[ref.ID]
		if h == nil || h.header == nil {
			continue
		}
		i++

		target := ""
		if rel := rels.ByID(ref.ID); rel != nil {
			target = rel.Target
		} else {
			target = fmt.Sprintf("header%d.xml", i)
			rels.AddWithID(ref.ID, packaging.RelTypeHeader, target, packaging.TargetModeInternal)
		}
		filename := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, target)

		data, err := utils.MarshalXMLWithHeader(h.header)
		if err != nil {
			return err
		}
		if _, err := d.pkg.AddPart(filename, packaging.ContentTypeHeader, data); err != nil {
			return err
		}
	}
	return nil
}

func (d *documentImpl) saveFooters() error {
	if d.document == nil || d.document.Body == nil || d.document.Body.SectPr == nil {
		return nil
	}

	rels := d.pkg.GetRelationships(packaging.WordDocumentPath)
	i := 0
	for _, ref := range d.document.Body.SectPr.FooterRefs {
		f := d.footers[ref.ID]
		if f == nil || f.footer == nil {
			continue
		}
		i++

		target := ""
		if rel := rels.ByID(ref.ID); rel != nil {
			target = rel.Target
		} else {
			target = fmt.Sprintf("footer%d.xml", i)
			rels.AddWithID(ref.ID, packaging.RelTypeFooter, target, packaging.TargetModeInternal)
		}
		filename := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, target)

		data, err := utils.MarshalXMLWithHeader(f.footer)
		if err != nil {
			return err
		}
		if _, err := d.pkg.AddPart(filename, packaging.ContentTypeFooter, data); err != nil {
			return err
		}
	}
	return nil
}

func (d *documentImpl) parseHeaders() error {
	if d.document == nil || d.document.Body == nil || d.document.Body.SectPr == nil {
		return nil
	}
	rels := d.pkg.GetRelationships(packaging.WordDocumentPath)
	for _, ref := range d.document.Body.SectPr.HeaderRefs {
		rel := rels.ByID(ref.ID)
		if rel == nil {
			continue
		}
		path := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, rel.Target)
		part, err := d.pkg.GetPart(path)
		if err != nil {
			continue
		}
		content, err := part.Content()
		if err != nil {
			return err
		}
		header := &wml.Header{}
		if err := utils.UnmarshalXML(content, header); err != nil {
			return err
		}
		d.headers[ref.ID] = &headerImpl{
			doc:    d,
			header: header,
			relID:  ref.ID,
			hfType: HeaderFooterType(ref.Type),
		}
	}
	return nil
}

func (d *documentImpl) parseFooters() error {
	if d.document == nil || d.document.Body == nil || d.document.Body.SectPr == nil {
		return nil
	}
	rels := d.pkg.GetRelationships(packaging.WordDocumentPath)
	for _, ref := range d.document.Body.SectPr.FooterRefs {
		rel := rels.ByID(ref.ID)
		if rel == nil {
			continue
		}
		path := packaging.ResolveRelationshipTarget(packaging.WordDocumentPath, rel.Target)
		part, err := d.pkg.GetPart(path)
		if err != nil {
			continue
		}
		content, err := part.Content()
		if err != nil {
			return err
		}
		footer := &wml.Footer{}
		if err := utils.UnmarshalXML(content, footer); err != nil {
			return err
		}
		d.footers[ref.ID] = &footerImpl{
			doc:    d,
			footer: footer,
			relID:  ref.ID,
			hfType: HeaderFooterType(ref.Type),
		}
	}
	return nil
}
