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

// Text returns the text inside the content control.
func (c *ContentControl) Text() string {
	return textFromSdt(c.sdt)
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
