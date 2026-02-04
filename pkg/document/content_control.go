// Package document provides content control functionality.
package document

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// ContentControl represents a content control (SDT).
type ContentControl struct {
	doc *documentImpl
	sdt *wml.Sdt
}

// ContentControlListItem represents a dropdown/combo box entry.
type ContentControlListItem struct {
	DisplayText string
	Value       string
}

// ContentControlDateConfig represents date picker properties.
type ContentControlDateConfig struct {
	Format           string
	Locale           string
	Calendar         string
	StoreMappedDataAs string
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

// ListItems returns dropdown/combobox items if configured.
func (c *ContentControl) ListItems() []ContentControlListItem {
	if c.sdt.SdtPr == nil {
		return nil
	}
	list := listItemsFromSdtPr(c.sdt.SdtPr)
	if list == nil {
		return nil
	}
	items := make([]ContentControlListItem, len(list.ListItem))
	for i, item := range list.ListItem {
		if item == nil {
			continue
		}
		items[i] = ContentControlListItem{
			DisplayText: item.DisplayText,
			Value:       item.Value,
		}
	}
	return items
}

// SetDropDownList configures the content control as a drop-down list.
func (c *ContentControl) SetDropDownList(items []ContentControlListItem) {
	c.ensureSdtPr()
	c.sdt.SdtPr.ComboBox = nil
	c.sdt.SdtPr.DropDownList = buildDropDownList(items)
}

// SetComboBox configures the content control as a combo box.
func (c *ContentControl) SetComboBox(items []ContentControlListItem) {
	c.ensureSdtPr()
	c.sdt.SdtPr.DropDownList = nil
	c.sdt.SdtPr.ComboBox = buildDropDownList(items)
}

// ClearListControl removes dropdown/combobox configuration.
func (c *ContentControl) ClearListControl() {
	if c.sdt.SdtPr == nil {
		return
	}
	c.sdt.SdtPr.DropDownList = nil
	c.sdt.SdtPr.ComboBox = nil
}

// DateConfig returns date picker configuration if set.
func (c *ContentControl) DateConfig() *ContentControlDateConfig {
	if c.sdt.SdtPr == nil || c.sdt.SdtPr.Date == nil {
		return nil
	}
	date := c.sdt.SdtPr.Date
	cfg := &ContentControlDateConfig{}
	if date.DateFormat != nil {
		cfg.Format = date.DateFormat.Val
	}
	if date.Language != nil {
		cfg.Locale = date.Language.Val
	}
	if date.Calendar != nil {
		cfg.Calendar = date.Calendar.Val
	}
	if date.StoreMappedDataAs != nil {
		cfg.StoreMappedDataAs = date.StoreMappedDataAs.Val
	}
	return cfg
}

// SetDateConfig configures the content control as a date picker.
func (c *ContentControl) SetDateConfig(cfg ContentControlDateConfig) {
	c.ensureSdtPr()
	date := &wml.SdtDate{}
	if cfg.Format != "" {
		date.DateFormat = &wml.SdtString{Val: cfg.Format}
	}
	if cfg.Locale != "" {
		date.Language = &wml.SdtLang{Val: cfg.Locale}
	}
	if cfg.Calendar != "" {
		date.Calendar = &wml.SdtString{Val: cfg.Calendar}
	}
	if cfg.StoreMappedDataAs != "" {
		date.StoreMappedDataAs = &wml.SdtString{Val: cfg.StoreMappedDataAs}
	}
	c.sdt.SdtPr.Date = date
}

// ClearDateConfig removes date picker configuration.
func (c *ContentControl) ClearDateConfig() {
	if c.sdt.SdtPr == nil {
		return
	}
	c.sdt.SdtPr.Date = nil
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
func (c *ContentControl) Paragraphs() []Paragraph {
	if c.sdt.SdtContent == nil {
		return nil
	}
	var result []Paragraph
	for i, elem := range c.sdt.SdtContent.Content {
		if p, ok := elem.(*wml.P); ok {
			result = append(result, &paragraphImpl{doc: c.doc, p: p, index: i})
		}
	}
	return result
}

// Tables returns the tables inside a block content control.
func (c *ContentControl) Tables() []Table {
	if c.sdt.SdtContent == nil {
		return nil
	}
	var result []Table
	for i, elem := range c.sdt.SdtContent.Content {
		if tbl, ok := elem.(*wml.Tbl); ok {
			result = append(result, &tableImpl{doc: c.doc, tbl: tbl, index: i})
		}
	}
	return result
}

// Runs returns the runs inside a run-level content control.
func (c *ContentControl) Runs() []Run {
	if c.sdt.SdtContent == nil {
		return nil
	}
	var result []Run
	for _, elem := range c.sdt.SdtContent.Content {
		if r, ok := elem.(*wml.R); ok {
			result = append(result, &runImpl{doc: c.doc, r: r})
		}
	}
	return result
}

// AddParagraph adds a paragraph to a block content control.
func (c *ContentControl) AddParagraph() Paragraph {
	c.ensureContent()
	p := &wml.P{}
	c.sdt.SdtContent.Content = append(c.sdt.SdtContent.Content, p)
	return &paragraphImpl{doc: c.doc, p: p, index: len(c.sdt.SdtContent.Content) - 1}
}

// AddRun adds a run to a run-level content control.
func (c *ContentControl) AddRun() Run {
	c.ensureContent()
	r := &wml.R{}
	c.sdt.SdtContent.Content = append(c.sdt.SdtContent.Content, r)
	return &runImpl{doc: c.doc, r: r}
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
		return utils.ErrContentControlNotFound
	}
	if removeSdtFromContent(&c.doc.document.Body.Content, c.sdt) {
		return nil
	}
	return utils.ErrContentControlNotFound
}

// AddContentControl adds a run-level content control to the paragraph.
func (p *paragraphImpl) AddContentControl(tag, alias, text string) *ContentControl {
	sdt := newContentControl(tag, alias, []interface{}{
		&wml.R{Content: []interface{}{wml.NewT(text)}},
	})
	p.p.Content = append(p.p.Content, sdt)
	return &ContentControl{doc: p.doc, sdt: sdt}
}

// AddBlockContentControl adds a block-level content control to the document body.
func (d *documentImpl) AddBlockContentControl(tag, alias, text string) *ContentControl {
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
func (d *documentImpl) ContentControls() []*ContentControl {
	if d.document == nil || d.document.Body == nil {
		return nil
	}
	var result []*ContentControl
	collectSdtFromContent(d.document.Body.Content, d, &result)
	return result
}

// AddContentControl adds an inline content control to the first paragraph.
func (d *documentImpl) AddContentControl(tag, alias, text string) *ContentControl {
	if d == nil {
		return nil
	}
	paras := d.Paragraphs()
	if len(paras) == 0 {
		return d.AddParagraph().AddContentControl(tag, alias, text)
	}
	return paras[0].AddContentControl(tag, alias, text)
}

// ContentControlsByTag returns all content controls with the specified tag.
func (d *documentImpl) ContentControlsByTag(tag string) []*ContentControl {
	var result []*ContentControl
	for _, cc := range d.ContentControls() {
		if cc.Tag() == tag {
			result = append(result, cc)
		}
	}
	return result
}

// ContentControlByTag returns the first content control with the specified tag.
func (d *documentImpl) ContentControlByTag(tag string) *ContentControl {
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
		return utils.NewValidationError("contentControlLock", "invalid value", lock)
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

func (c *ContentControl) ensureSdtPr() {
	if c.sdt.SdtPr == nil {
		c.sdt.SdtPr = &wml.SdtPr{}
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

func collectSdtFromContent(content []interface{}, doc *documentImpl, result *[]*ContentControl) {
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

func listItemsFromSdtPr(pr *wml.SdtPr) *wml.SdtDropDownList {
	if pr == nil {
		return nil
	}
	if pr.DropDownList != nil {
		return pr.DropDownList
	}
	return pr.ComboBox
}

func buildDropDownList(items []ContentControlListItem) *wml.SdtDropDownList {
	list := &wml.SdtDropDownList{}
	for _, item := range items {
		listItem := &wml.SdtListItem{}
		if item.DisplayText != "" {
			listItem.DisplayText = item.DisplayText
		}
		if item.Value != "" {
			listItem.Value = item.Value
		}
		list.ListItem = append(list.ListItem, listItem)
	}
	return list
}
