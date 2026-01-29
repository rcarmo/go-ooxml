package document

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

// Paragraph represents a paragraph in a Word document.
type Paragraph struct {
	doc   *Document
	p     *wml.P
	index int
}

// Text returns the combined text of all runs in the paragraph.
func (p *Paragraph) Text() string {
	var sb strings.Builder
	for _, run := range p.Runs() {
		sb.WriteString(run.Text())
	}
	return sb.String()
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

// Index returns the paragraph's index in the body.
func (p *Paragraph) Index() int {
	return p.index
}

// XML returns the underlying WML paragraph for advanced access.
func (p *Paragraph) XML() *wml.P {
	return p.p
}
