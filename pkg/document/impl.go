package document

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

type bodyImpl struct {
	doc *documentImpl
}

type paragraphImpl struct {
	doc   *documentImpl
	p     *wml.P
	index int
}

type runImpl struct {
	doc *documentImpl
	r   *wml.R
}

type tableImpl struct {
	doc   *documentImpl
	tbl   *wml.Tbl
	index int
}

type rowImpl struct {
	doc   *documentImpl
	tr    *wml.Tr
	index int
}

type cellImpl struct {
	doc   *documentImpl
	tc    *wml.Tc
	index int
}

type headerImpl struct {
	doc    *documentImpl
	header *wml.Header
	relID  string
	hfType HeaderFooterType
}

type footerImpl struct {
	doc    *documentImpl
	footer *wml.Footer
	relID  string
	hfType HeaderFooterType
}

type stylesImpl struct {
	doc *documentImpl
}

type styleImpl struct {
	doc   *documentImpl
	style *wml.Style
}

type commentsImpl struct {
	doc *documentImpl
}

type commentImpl struct {
	doc     *documentImpl
	comment *wml.Comment
	paraID  string
}

type sectionImpl struct {
	doc    *documentImpl
	sectPr *wml.SectPr
}

// Header returns the header for the section.
func (s *sectionImpl) Header(hfType HeaderFooterType) Header {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.Header(hfType)
}

// Footer returns the footer for the section.
func (s *sectionImpl) Footer(hfType HeaderFooterType) Footer {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.Footer(hfType)
}

// AddHeader adds a header for the section.
func (s *sectionImpl) AddHeader(hfType HeaderFooterType) Header {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.AddHeader(hfType)
}

// AddFooter adds a footer for the section.
func (s *sectionImpl) AddFooter(hfType HeaderFooterType) Footer {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.AddFooter(hfType)
}

// PageMargins returns the page margins for the section.
func (s *sectionImpl) PageMargins() (PageMargins, bool) {
	if s == nil || s.sectPr == nil || s.sectPr.PgMar == nil {
		return PageMargins{}, false
	}
	return *s.sectPr.PgMar, true
}

// SetPageMargins sets the page margins for the section.
func (s *sectionImpl) SetPageMargins(margins PageMargins) {
	if s == nil {
		return
	}
	if s.sectPr == nil {
		s.sectPr = &wml.SectPr{}
	}
	s.sectPr.PgMar = &margins
}

// TitlePage reports whether the section uses a different first page.
func (s *sectionImpl) TitlePage() bool {
	if s == nil || s.sectPr == nil || s.sectPr.TitlePg == nil {
		return false
	}
	return s.sectPr.TitlePg.Enabled()
}

// SetTitlePage sets whether the section uses a different first page.
func (s *sectionImpl) SetTitlePage(v bool) {
	if s == nil {
		return
	}
	if s.sectPr == nil {
		s.sectPr = &wml.SectPr{}
	}
	if v {
		s.sectPr.TitlePg = wml.NewOnOffEnabled()
	} else {
		s.sectPr.TitlePg = nil
	}
}
