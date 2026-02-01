// Package document provides content control functionality.
package document

import (
	"fmt"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// ContentControl represents a content control (SDT).
type ContentControl struct {
	doc *Document
	sdt *wml.Sdt
}

// Tag returns the content control tag.
func (c *ContentControl) Tag() string {
	if c.sdt.SdtPr != nil && c.sdt.SdtPr.Tag != nil {
		return c.sdt.SdtPr.Tag.Val
	}
	return ""
}

// Alias returns the content control alias.
func (c *ContentControl) Alias() string {
	if c.sdt.SdtPr != nil && c.sdt.SdtPr.Alias != nil {
		return c.sdt.SdtPr.Alias.Val
	}
	return ""
}

// SetTag sets the content control tag.
func (c *ContentControl) SetTag(tag string) {
	if c.sdt.SdtPr == nil {
		c.sdt.SdtPr = &wml.SdtPr{}
	}
	if tag == "" {
		c.sdt.SdtPr.Tag = nil
		return
	}
	c.sdt.SdtPr.Tag = &wml.SdtString{Val: tag}
}

// SetAlias sets the content control alias.
func (c *ContentControl) SetAlias(alias string) {
	if c.sdt.SdtPr == nil {
		c.sdt.SdtPr = &wml.SdtPr{}
	}
	if alias == "" {
		c.sdt.SdtPr.Alias = nil
		return
	}
	c.sdt.SdtPr.Alias = &wml.SdtString{Val: alias}
}

// ID returns the content control ID.
func (c *ContentControl) ID() int {
	if c.sdt.SdtPr != nil && c.sdt.SdtPr.ID != nil {
		return c.sdt.SdtPr.ID.Val
	}
	return 0
}

// Lock returns the content control lock setting.
func (c *ContentControl) Lock() string {
	if c.sdt.SdtPr != nil && c.sdt.SdtPr.Lock != nil {
		return c.sdt.SdtPr.Lock.Val
	}
	return ""
}

// Text returns the text inside the content control.
func (c *ContentControl) Text() string {
	return textFromSdt(c.sdt)
}

// SetText replaces the content control text.
func (c *ContentControl) SetText(text string) {
	c.ensureContent()
	if c.IsInline() {
		c.sdt.SdtContent.Content = []interface{}{
			&wml.R{Content: []interface{}{wml.NewT(text)}},
		}
		return
	}
	c.sdt.SdtContent.Content = []interface{}{
		&wml.P{Content: []interface{}{&wml.R{Content: []interface{}{wml.NewT(text)}}}},
	}
}

// Paragraphs returns the paragraphs inside a block content control.
func (c *ContentControl) Paragraphs() []*Paragraph {
	if c.sdt.SdtContent == nil {
		return nil
	}
	var result []*Paragraph
	for i, elem := range c.sdt.SdtContent.Content {
		if p, ok := elem.(*wml.P); ok {
			result = append(result, &Paragraph{doc: c.doc, p: p, index: i})
		}
	}
	return result
}

// Tables returns the tables inside a block content control.
func (c *ContentControl) Tables() []*Table {
	if c.sdt.SdtContent == nil {
		return nil
	}
	var result []*Table
	for i, elem := range c.sdt.SdtContent.Content {
		if tbl, ok := elem.(*wml.Tbl); ok {
			result = append(result, &Table{doc: c.doc, tbl: tbl, index: i})
		}
	}
	return result
}

// Runs returns the runs inside a run-level content control.
func (c *ContentControl) Runs() []*Run {
	if c.sdt.SdtContent == nil {
		return nil
	}
	var result []*Run
	for _, elem := range c.sdt.SdtContent.Content {
		if r, ok := elem.(*wml.R); ok {
			result = append(result, &Run{doc: c.doc, r: r})
		}
	}
	return result
}

// AddParagraph adds a paragraph to a block content control.
func (c *ContentControl) AddParagraph() *Paragraph {
	c.ensureContent()
	p := &wml.P{}
	c.sdt.SdtContent.Content = append(c.sdt.SdtContent.Content, p)
	return &Paragraph{doc: c.doc, p: p, index: len(c.sdt.SdtContent.Content) - 1}
}

// AddRun adds a run to a run-level content control.
func (c *ContentControl) AddRun() *Run {
	c.ensureContent()
	r := &wml.R{}
	c.sdt.SdtContent.Content = append(c.sdt.SdtContent.Content, r)
	return &Run{doc: c.doc, r: r}
}

// IsInline returns true for run-level content controls.
func (c *ContentControl) IsInline() bool {
	if c.doc != nil && c.doc.document != nil && c.doc.document.Body != nil {
		found, inline := findSdtInContent(c.doc.document.Body.Content, c.sdt, false)
		if found {
			return inline
		}
	}
	return hasInlineContent(c.sdt)
}

// IsBlock returns true for block-level content controls.
func (c *ContentControl) IsBlock() bool {
	return !c.IsInline()
}

// Remove removes the content control from the document.
func (c *ContentControl) Remove() error {
	if c.doc == nil || c.doc.document == nil || c.doc.document.Body == nil || c.sdt == nil {
		return fmt.Errorf("content control not found")
	}
	if removeSdtFromContent(&c.doc.document.Body.Content, c.sdt) {
		return nil
	}
	return fmt.Errorf("content control not found")
}

// AddContentControl adds a run-level content control to the paragraph.
func (p *Paragraph) AddContentControl(tag, alias, text string) *ContentControl {
	sdt := newContentControl(tag, alias, []interface{}{
		&wml.R{Content: []interface{}{wml.NewT(text)}},
	})
	p.p.Content = append(p.p.Content, sdt)
	return &ContentControl{doc: p.doc, sdt: sdt}
}

// AddBlockContentControl adds a block-level content control to the document body.
func (d *Document) AddBlockContentControl(tag, alias, text string) *ContentControl {
	p := &wml.P{Content: []interface{}{&wml.R{Content: []interface{}{wml.NewT(text)}}}}
	sdt := newContentControl(tag, alias, []interface{}{p})
	if d.document.Body == nil {
		d.document.Body = &wml.Body{}
	}
	d.document.Body.Content = append(d.document.Body.Content, sdt)
	return &ContentControl{doc: d, sdt: sdt}
}

func newContentControl(tag, alias string, content []interface{}) *wml.Sdt {
	var sdtPr *wml.SdtPr
	if tag != "" || alias != "" {
		sdtPr = &wml.SdtPr{}
		if tag != "" {
			sdtPr.Tag = &wml.SdtString{Val: tag}
		}
		if alias != "" {
			sdtPr.Alias = &wml.SdtString{Val: alias}
		}
	}
	return &wml.Sdt{
		SdtPr:      sdtPr,
		SdtContent: &wml.SdtContent{Content: content},
	}
}

// ContentControls returns all content controls in the document.
func (d *Document) ContentControls() []*ContentControl {
	if d.document == nil || d.document.Body == nil {
		return nil
	}
	var result []*ContentControl
	collectSdtFromContent(d.document.Body.Content, d, &result)
	return result
}

// ContentControlsByTag returns all content controls with the specified tag.
func (d *Document) ContentControlsByTag(tag string) []*ContentControl {
	var result []*ContentControl
	for _, cc := range d.ContentControls() {
		if cc.Tag() == tag {
			result = append(result, cc)
		}
	}
	return result
}

// ContentControlByTag returns the first content control with the specified tag.
func (d *Document) ContentControlByTag(tag string) *ContentControl {
	for _, cc := range d.ContentControls() {
		if cc.Tag() == tag {
			return cc
		}
	}
	return nil
}

// SetContentControlID sets the SDT ID on a content control.
func (c *ContentControl) SetContentControlID(id int) {
	if id <= 0 {
		return
	}
	if c.sdt.SdtPr == nil {
		c.sdt.SdtPr = &wml.SdtPr{}
	}
	c.sdt.SdtPr.ID = &wml.SdtID{Val: id}
}

// SetContentControlLock sets a lock value on the content control.
func (c *ContentControl) SetContentControlLock(lock string) error {
	if lock == "" {
		if c.sdt.SdtPr != nil {
			c.sdt.SdtPr.Lock = nil
		}
		return nil
	}
	switch lock {
	case "sdt", "contentControl", "content":
		if c.sdt.SdtPr == nil {
			c.sdt.SdtPr = &wml.SdtPr{}
		}
		c.sdt.SdtPr.Lock = &wml.SdtLock{Val: lock}
		return nil
	default:
		return fmt.Errorf("invalid content control lock: %s", lock)
	}
}

// XML returns the underlying WML content control.
func (c *ContentControl) XML() *wml.Sdt {
	return c.sdt
}

func (c *ContentControl) ensureContent() {
	if c.sdt.SdtContent == nil {
		c.sdt.SdtContent = &wml.SdtContent{}
	}
}

func hasInlineContent(sdt *wml.Sdt) bool {
	if sdt == nil || sdt.SdtContent == nil {
		return false
	}
	for _, elem := range sdt.SdtContent.Content {
		switch elem.(type) {
		case *wml.R, *wml.Hyperlink:
			return true
		case *wml.P, *wml.Tbl:
			return false
		}
	}
	return false
}

func findSdtInContent(content []interface{}, target *wml.Sdt, inlineContext bool) (bool, bool) {
	for _, elem := range content {
		switch v := elem.(type) {
		case *wml.Sdt:
			if v == target {
				return true, inlineContext
			}
			if v.SdtContent != nil {
				found, inline := findSdtInContent(v.SdtContent.Content, target, inlineContext)
				if found {
					return found, inline
				}
			}
		case *wml.P:
			found, inline := findSdtInContent(v.Content, target, true)
			if found {
				return found, inline
			}
		case *wml.Tbl:
			found, inline := findSdtInTable(v, target)
			if found {
				return found, inline
			}
		case *wml.Hyperlink:
			found, inline := findSdtInContent(v.Content, target, true)
			if found {
				return found, inline
			}
		}
	}
	return false, false
}

func findSdtInTable(tbl *wml.Tbl, target *wml.Sdt) (bool, bool) {
	for _, row := range tbl.Tr {
		for _, cell := range row.Tc {
			found, inline := findSdtInContent(cell.Content, target, false)
			if found {
				return found, inline
			}
		}
	}
	return false, false
}

func collectSdtFromContent(content []interface{}, doc *Document, result *[]*ContentControl) {
	for _, elem := range content {
		switch v := elem.(type) {
		case *wml.Sdt:
			*result = append(*result, &ContentControl{doc: doc, sdt: v})
			if v.SdtContent != nil {
				collectSdtFromContent(v.SdtContent.Content, doc, result)
			}
		case *wml.P:
			collectSdtFromContent(v.Content, doc, result)
		case *wml.Tbl:
			for _, row := range v.Tr {
				for _, cell := range row.Tc {
					collectSdtFromContent(cell.Content, doc, result)
				}
			}
		case *wml.Hyperlink:
			collectSdtFromContent(v.Content, doc, result)
		}
	}
}

func removeSdtFromContent(content *[]interface{}, target *wml.Sdt) bool {
	if content == nil {
		return false
	}
	for i := 0; i < len(*content); i++ {
		elem := (*content)[i]
		switch v := elem.(type) {
		case *wml.Sdt:
			if v == target {
				*content = append((*content)[:i], (*content)[i+1:]...)
				return true
			}
			if v.SdtContent != nil && removeSdtFromContent(&v.SdtContent.Content, target) {
				return true
			}
		case *wml.P:
			if removeSdtFromContent(&v.Content, target) {
				return true
			}
		case *wml.Tbl:
			if removeSdtFromTable(v, target) {
				return true
			}
		case *wml.Hyperlink:
			if removeSdtFromContent(&v.Content, target) {
				return true
			}
		}
	}
	return false
}

func removeSdtFromTable(tbl *wml.Tbl, target *wml.Sdt) bool {
	for _, row := range tbl.Tr {
		for _, cell := range row.Tc {
			if removeSdtFromContent(&cell.Content, target) {
				return true
			}
		}
	}
	return false
}
