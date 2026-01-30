package document

import (
	"fmt"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Paragraph represents a paragraph in a Word document.
type Paragraph struct {
	doc   *Document
	p     *wml.P
	index int
}

// Text returns the combined text of all runs in the paragraph.
func (p *Paragraph) Text() string {
	return textFromParagraph(p.p)
}

// SetText sets the paragraph text, replacing all existing runs.
func (p *Paragraph) SetText(text string) {
	// Clear existing content
	p.p.Content = nil

	// Add a single run with the text
	run := p.AddRun()
	run.SetText(text)
}

// Runs returns all runs in the paragraph.
func (p *Paragraph) Runs() []*Run {
	var result []*Run
	for _, elem := range p.p.Content {
		if r, ok := elem.(*wml.R); ok {
			result = append(result, &Run{doc: p.doc, r: r})
		}
	}
	return result
}

// AddBookmark inserts a bookmark start/end around the specified run range.
func (p *Paragraph) AddBookmark(name string, startRun, endRun int) error {
	if name == "" {
		return fmt.Errorf("bookmark name cannot be empty")
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
		return fmt.Errorf("document not available")
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
func (p *Paragraph) Hyperlinks() []*Hyperlink {
	var result []*Hyperlink
	for _, elem := range p.p.Content {
		if h, ok := elem.(*wml.Hyperlink); ok {
			result = append(result, &Hyperlink{doc: p.doc, h: h})
		}
	}
	return result
}

// ContentControls returns all content controls in the paragraph.
func (p *Paragraph) ContentControls() []*ContentControl {
	var result []*ContentControl
	for _, elem := range p.p.Content {
		if sdt, ok := elem.(*wml.Sdt); ok {
			result = append(result, &ContentControl{doc: p.doc, sdt: sdt})
		}
	}
	return result
}

// AddRun adds a new run to the paragraph.
func (p *Paragraph) AddRun() *Run {
	r := &wml.R{}
	p.p.Content = append(p.p.Content, r)
	return &Run{doc: p.doc, r: r}
}

// Style returns the paragraph style ID.
func (p *Paragraph) Style() string {
	if p.p.PPr != nil && p.p.PPr.PStyle != nil {
		return p.p.PPr.PStyle.Val
	}
	return ""
}

// SetStyle sets the paragraph style.
func (p *Paragraph) SetStyle(styleID string) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	p.p.PPr.PStyle = &wml.PStyle{Val: styleID}
}

// IsHeading returns whether the paragraph is a heading.
func (p *Paragraph) IsHeading() bool {
	style := p.Style()
	return strings.HasPrefix(style, "Heading") || strings.HasPrefix(style, "heading")
}

// HeadingLevel returns the heading level (1-9) or 0 if not a heading.
func (p *Paragraph) HeadingLevel() int {
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
func (p *Paragraph) Alignment() string {
	if p.p.PPr != nil && p.p.PPr.Jc != nil {
		return p.p.PPr.Jc.Val
	}
	return ""
}

// SetAlignment sets the paragraph alignment (left, center, right, both).
func (p *Paragraph) SetAlignment(align string) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	p.p.PPr.Jc = &wml.Jc{Val: align}
}

// SpacingBefore returns the spacing before the paragraph in twips.
func (p *Paragraph) SpacingBefore() int64 {
	if p.p.PPr != nil && p.p.PPr.Spacing != nil && p.p.PPr.Spacing.Before != nil {
		return *p.p.PPr.Spacing.Before
	}
	return 0
}

// SetSpacingBefore sets the spacing before the paragraph in twips.
func (p *Paragraph) SetSpacingBefore(twips int64) {
	p.ensureSpacing()
	p.p.PPr.Spacing.Before = &twips
}

// SpacingAfter returns the spacing after the paragraph in twips.
func (p *Paragraph) SpacingAfter() int64 {
	if p.p.PPr != nil && p.p.PPr.Spacing != nil && p.p.PPr.Spacing.After != nil {
		return *p.p.PPr.Spacing.After
	}
	return 0
}

// SetSpacingAfter sets the spacing after the paragraph in twips.
func (p *Paragraph) SetSpacingAfter(twips int64) {
	p.ensureSpacing()
	p.p.PPr.Spacing.After = &twips
}

func (p *Paragraph) ensureSpacing() {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if p.p.PPr.Spacing == nil {
		p.p.PPr.Spacing = &wml.Spacing{}
	}
}

// KeepWithNext returns whether the paragraph is kept with the next.
func (p *Paragraph) KeepWithNext() bool {
	if p.p.PPr != nil && p.p.PPr.KeepNext != nil {
		return p.p.PPr.KeepNext.Enabled()
	}
	return false
}

// SetKeepWithNext sets whether to keep with the next paragraph.
func (p *Paragraph) SetKeepWithNext(v bool) {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if v {
		p.p.PPr.KeepNext = wml.NewOnOffEnabled()
	} else {
		p.p.PPr.KeepNext = nil
	}
}

// ListLevel returns the list level for the paragraph or -1 if not a list item.
func (p *Paragraph) ListLevel() int {
	if p.p.PPr == nil || p.p.PPr.NumPr == nil || p.p.PPr.NumPr.Ilvl == nil {
		return -1
	}
	return p.p.PPr.NumPr.Ilvl.Val
}

// SetListLevel sets the list level (0-8) for the paragraph.
func (p *Paragraph) SetListLevel(level int) error {
	if level < 0 || level > 8 {
		return utils.ErrInvalidIndex
	}
	p.ensureNumPr()
	p.p.PPr.NumPr.Ilvl = &wml.Ilvl{Val: level}
	return nil
}

// ListNumberingID returns the numbering definition ID or 0 if not set.
func (p *Paragraph) ListNumberingID() int {
	if p.p.PPr == nil || p.p.PPr.NumPr == nil || p.p.PPr.NumPr.NumID == nil {
		return 0
	}
	return p.p.PPr.NumPr.NumID.Val
}

// SetListNumberingID sets the numbering definition ID for the paragraph.
func (p *Paragraph) SetListNumberingID(numID int) {
	p.ensureNumPr()
	p.p.PPr.NumPr.NumID = &wml.NumID{Val: numID}
}

// SetList sets the list numbering ID and level on the paragraph.
func (p *Paragraph) SetList(numID, level int) error {
	if err := p.SetListLevel(level); err != nil {
		return err
	}
	p.SetListNumberingID(numID)
	return nil
}

func (p *Paragraph) ensureNumPr() {
	if p.p.PPr == nil {
		p.p.PPr = &wml.PPr{}
	}
	if p.p.PPr.NumPr == nil {
		p.p.PPr.NumPr = &wml.NumPr{}
	}
}

// Index returns the paragraph's index in the body.
func (p *Paragraph) Index() int {
	return p.index
}

// XML returns the underlying WML paragraph for advanced access.
func (p *Paragraph) XML() *wml.P {
	return p.p
}
