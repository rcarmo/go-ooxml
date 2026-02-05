package document

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Text returns the combined text of all runs in the paragraph.
func (p *paragraphImpl) Text() string {
	return textFromParagraph(p.p)
}

// SetText sets the paragraph text, replacing all existing runs.
func (p *paragraphImpl) SetText(text string) {
	// Clear existing content
	p.p.Content = nil

	// Add a single run with the text
	run := p.AddRun()
	run.SetText(text)
}

// Runs returns all runs in the paragraph.
func (p *paragraphImpl) Runs() []Run {
	var result []Run
	for _, elem := range p.p.Content {
		if r, ok := elem.(*wml.R); ok {
			result = append(result, &runImpl{doc: p.doc, r: r})
		}
	}
	return result
}

// AddBookmark inserts a bookmark start/end around the specified run range.
func (p *paragraphImpl) AddBookmark(name string, startRun, endRun int) error {
	if name == "" {
		return utils.NewValidationError("bookmark", "name cannot be empty", name)
	}
	if len(p.p.Content) == 0 {
		return utils.ErrInvalidIndex
	}
	if startRun < 0 || endRun < startRun {
		return utils.ErrInvalidIndex
	}
	runCount := 0
	for _, elem := range p.p.Content {
		if _, ok := elem.(*wml.R); ok {
			runCount++
		}
	}
	if runCount == 0 || startRun >= runCount || endRun >= runCount {
		return utils.ErrInvalidIndex
	}
	if p.doc == nil {
		return utils.ErrDocumentClosed
	}

	id := p.doc.nextBookmarkID
	p.doc.nextBookmarkID++
	start := &wml.BookmarkStart{ID: id, Name: name}
	end := &wml.BookmarkEnd{ID: id}

	newContent := make([]interface{}, 0, len(p.p.Content)+2)
	runIndex := 0
	for _, elem := range p.p.Content {
		if _, ok := elem.(*wml.R); ok {
			if runIndex == startRun {
				newContent = append(newContent, start)
			}
			runIndex++
		}
		newContent = append(newContent, elem)
		if _, ok := elem.(*wml.R); ok && runIndex-1 == endRun {
			newContent = append(newContent, end)
		}
	}
	p.p.Content = newContent
	return nil
}

// Hyperlinks returns all hyperlinks in the paragraph.
func (p *paragraphImpl) Hyperlinks() []*Hyperlink {
	var result []*Hyperlink
	for _, elem := range p.p.Content {
		if h, ok := elem.(*wml.Hyperlink); ok {
			result = append(result, &Hyperlink{doc: p.doc, h: h})
		}
	}
	return result
}

// ContentControls returns all content controls in the paragraph.
func (p *paragraphImpl) ContentControls() []*ContentControl {
	var result []*ContentControl
	for _, elem := range p.p.Content {
		if sdt, ok := elem.(*wml.Sdt); ok {
			result = append(result, &ContentControl{doc: p.doc, sdt: sdt})
		}
	}
	return result
}

// AddRun adds a new run to the paragraph.
func (p *paragraphImpl) AddRun() Run {
	r := &wml.R{}
	p.p.Content = append(p.p.Content, r)
	return &runImpl{doc: p.doc, r: r}
}

// Style returns the paragraph style ID.
func (p *paragraphImpl) Style() string {
	if p.p.PPr != nil && p.p.PPr.PStyle != nil {
		return p.p.PPr.PStyle.Val
	}
	return ""
}

// SetStyle sets the paragraph style.
func (p *paragraphImpl) SetStyle(styleID string) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	p.p.PPr.PStyle = &wml.PStyle{Val: styleID}
}

// Properties returns the paragraph properties.
func (p *paragraphImpl) Properties() ParagraphProperties {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	return *p.p.PPr
}

// IsHeading returns whether the paragraph is a heading.
func (p *paragraphImpl) IsHeading() bool {
	style := p.Style()
	return strings.HasPrefix(style, "Heading") || strings.HasPrefix(style, "heading")
}

// HeadingLevel returns the heading level (1-9) or 0 if not a heading.
func (p *paragraphImpl) HeadingLevel() int {
	style := p.Style()
	if strings.HasPrefix(style, "Heading") {
		if len(style) > 7 {
			level := int(style[7] - '0')
			if level >= 1 && level <= 9 {
				return level
			}
		}
	}
	// Check outline level
	if p.p.PPr != nil && p.p.PPr.OutlineLvl != nil {
		return p.p.PPr.OutlineLvl.Val + 1 // outlineLvl is 0-based
	}
	return 0
}

// Alignment returns the paragraph alignment.
func (p *paragraphImpl) Alignment() string {
	if p.p.PPr != nil && p.p.PPr.Jc != nil {
		return p.p.PPr.Jc.Val
	}
	return ""
}

// SetAlignment sets the paragraph alignment (left, center, right, both).
func (p *paragraphImpl) SetAlignment(align string) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	p.p.PPr.Jc = &wml.Jc{Val: align}
}

// SpacingBefore returns the spacing before the paragraph in twips.
func (p *paragraphImpl) SpacingBefore() int64 {
	if p.p.PPr != nil && p.p.PPr.Spacing != nil && p.p.PPr.Spacing.Before != nil {
		return *p.p.PPr.Spacing.Before
	}
	return 0
}

// SetSpacingBefore sets the spacing before the paragraph in twips.
func (p *paragraphImpl) SetSpacingBefore(twips int64) {
	p.ensureSpacing()
	p.p.PPr.Spacing.Before = &twips
}

// SpacingAfter returns the spacing after the paragraph in twips.
func (p *paragraphImpl) SpacingAfter() int64 {
	if p.p.PPr != nil && p.p.PPr.Spacing != nil && p.p.PPr.Spacing.After != nil {
		return *p.p.PPr.Spacing.After
	}
	return 0
}

// SetSpacingAfter sets the spacing after the paragraph in twips.
func (p *paragraphImpl) SetSpacingAfter(twips int64) {
	p.ensureSpacing()
	p.p.PPr.Spacing.After = &twips
}

func (p *paragraphImpl) ensureSpacing() {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if p.p.PPr.Spacing == nil {
		p.p.PPr.Spacing = &wml.Spacing{}
	}
}

// KeepWithNext returns whether the paragraph is kept with the next.
func (p *paragraphImpl) KeepWithNext() bool {
	if p.p.PPr != nil && p.p.PPr.KeepNext != nil {
		return p.p.PPr.KeepNext.Enabled()
	}
	return false
}

// SetKeepWithNext sets whether to keep with the next paragraph.
func (p *paragraphImpl) SetKeepWithNext(v bool) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if v {
		p.p.PPr.KeepNext = wml.NewOnOffEnabled()
	} else {
		p.p.PPr.KeepNext = nil
	}
}

// KeepLines returns whether the paragraph keeps lines together.
func (p *paragraphImpl) KeepLines() bool {
	if p.p.PPr != nil && p.p.PPr.KeepLines != nil {
		return p.p.PPr.KeepLines.Enabled()
	}
	return false
}

// SetKeepLines sets whether to keep paragraph lines together.
func (p *paragraphImpl) SetKeepLines(v bool) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if v {
		p.p.PPr.KeepLines = wml.NewOnOffEnabled()
	} else {
		p.p.PPr.KeepLines = nil
	}
}

// PageBreakBefore returns whether the paragraph has a page break before.
func (p *paragraphImpl) PageBreakBefore() bool {
	if p.p.PPr != nil && p.p.PPr.PageBreakBefore != nil {
		return p.p.PPr.PageBreakBefore.Enabled()
	}
	return false
}

// SetPageBreakBefore sets whether to insert a page break before the paragraph.
func (p *paragraphImpl) SetPageBreakBefore(v bool) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if v {
		p.p.PPr.PageBreakBefore = wml.NewOnOffEnabled()
	} else {
		p.p.PPr.PageBreakBefore = nil
	}
}

// WidowControl returns whether widow/orphan control is enabled.
func (p *paragraphImpl) WidowControl() bool {
	if p.p.PPr != nil && p.p.PPr.WidowControl != nil {
		return p.p.PPr.WidowControl.Enabled()
	}
	return false
}

// SetWidowControl sets widow/orphan control.
func (p *paragraphImpl) SetWidowControl(v bool) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if v {
		p.p.PPr.WidowControl = wml.NewOnOffEnabled()
	} else {
		p.p.PPr.WidowControl = nil
	}
}

// ListLevel returns the list level for the paragraph or -1 if not a list item.
func (p *paragraphImpl) ListLevel() int {
	if p.p.PPr == nil || p.p.PPr.NumPr == nil || p.p.PPr.NumPr.Ilvl == nil {
		return -1
	}
	return p.p.PPr.NumPr.Ilvl.Val
}

// SetListLevel sets the list level (0-8) for the paragraph.
func (p *paragraphImpl) SetListLevel(level int) error {
	if level < 0 || level > 8 {
		return utils.ErrInvalidIndex
	}
	p.ensureNumPr()
	p.p.PPr.NumPr.Ilvl = &wml.Ilvl{Val: level}
	return nil
}

// ListNumberingID returns the numbering definition ID or 0 if not set.
func (p *paragraphImpl) ListNumberingID() int {
	if p.p.PPr == nil || p.p.PPr.NumPr == nil || p.p.PPr.NumPr.NumID == nil {
		return 0
	}
	return p.p.PPr.NumPr.NumID.Val
}

// SetListNumberingID sets the numbering definition ID for the paragraph.
func (p *paragraphImpl) SetListNumberingID(numID int) {
	p.ensureNumPr()
	p.p.PPr.NumPr.NumID = &wml.NumID{Val: numID}
}

// SetList sets the list numbering ID and level on the paragraph.
func (p *paragraphImpl) SetList(numID, level int) error {
	if err := p.SetListLevel(level); err != nil {
		return err
	}
	p.SetListNumberingID(numID)
	return nil
}

func (p *paragraphImpl) ensureNumPr() {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if p.p.PPr.NumPr == nil {
		p.p.PPr.NumPr = &wml.NumPr{}
	}
}

// Index returns the paragraph's index in the body.
func (p *paragraphImpl) Index() int {
	return p.index
}

// XML returns the underlying WML paragraph for advanced access.
func (p *paragraphImpl) XML() *wml.P {
	return p.p
}
