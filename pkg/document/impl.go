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
}

type sectionImpl struct {
	doc    *documentImpl
	sectPr *wml.SectPr
}

func (s *sectionImpl) Header(hfType HeaderFooterType) Header {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.Header(hfType)
}

func (s *sectionImpl) Footer(hfType HeaderFooterType) Footer {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.Footer(hfType)
}

func (s *sectionImpl) AddHeader(hfType HeaderFooterType) Header {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.AddHeader(hfType)
}

func (s *sectionImpl) AddFooter(hfType HeaderFooterType) Footer {
	if s == nil || s.doc == nil {
		return nil
	}
	return s.doc.AddFooter(hfType)
}
