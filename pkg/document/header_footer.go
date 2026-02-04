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
	HeaderFooterDefault HeaderFooterType = "default"
	HeaderFooterFirst   HeaderFooterType = "first"
	HeaderFooterEven    HeaderFooterType = "even"
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
	// Generate unique filename
	num := len(d.headers) + 1
	filename := fmt.Sprintf("header%d.xml", num)
	relID := fmt.Sprintf("rId%d", 100+num)
	
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
	d.pkg.AddRelationship(packaging.WordDocumentPath, filename, packaging.RelTypeHeader)
	
	return h
}

// AddFooter adds a footer of the specified type.
func (d *documentImpl) AddFooter(hfType HeaderFooterType) Footer {
	// Generate unique filename
	num := len(d.footers) + 1
	filename := fmt.Sprintf("footer%d.xml", num)
	relID := fmt.Sprintf("rId%d", 200+num)
	
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
	d.pkg.AddRelationship(packaging.WordDocumentPath, filename, packaging.RelTypeFooter)
	
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
	i := 0
	for _, h := range d.headers {
		if h.header == nil {
			continue
		}
		i++
		
		data, err := utils.MarshalXMLWithHeader(h.header)
		if err != nil {
			return err
		}
		
		filename := fmt.Sprintf("word/header%d.xml", i)
		_, err = d.pkg.AddPart(filename, packaging.ContentTypeHeader, data)
		if err != nil {
			return err
		}
	}
	return nil
}

func (d *documentImpl) saveFooters() error {
	i := 0
	for _, f := range d.footers {
		if f.footer == nil {
			continue
		}
		i++
		
		data, err := utils.MarshalXMLWithHeader(f.footer)
		if err != nil {
			return err
		}
		
		filename := fmt.Sprintf("word/footer%d.xml", i)
		_, err = d.pkg.AddPart(filename, packaging.ContentTypeFooter, data)
		if err != nil {
			return err
		}
	}
	return nil
}
