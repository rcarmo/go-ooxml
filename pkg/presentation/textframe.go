package presentation

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
)

// TextFrame represents the text content of a shape.
type TextFrame struct {
	shape  *Shape
	txBody *dml.TxBody
}

// Text returns all text in the text frame.
func (tf *TextFrame) Text() string {
	return tf.shape.Text()
}

// SetText sets all text in the text frame.
func (tf *TextFrame) SetText(text string) {
	tf.shape.SetText(text)
}

// Paragraphs returns all paragraphs in the text frame.
func (tf *TextFrame) Paragraphs() []*TextParagraph {
	if tf.txBody == nil || len(tf.txBody.P) == 0 {
		return nil
	}

	var result []*TextParagraph
	for i, p := range tf.txBody.P {
		result = append(result, &TextParagraph{
			textFrame: tf,
			p:         p,
			index:     i,
		})
	}
	return result
}

// AddParagraph adds a new paragraph to the text frame.
func (tf *TextFrame) AddParagraph() *TextParagraph {
	p := &dml.P{}
	tf.txBody.P = append(tf.txBody.P, p)
	return &TextParagraph{
		textFrame: tf,
		p:         p,
		index:     len(tf.txBody.P) - 1,
	}
}

// ClearParagraphs removes all paragraphs from the text frame.
func (tf *TextFrame) ClearParagraphs() {
	tf.txBody.P = nil
}

// AutofitType returns the autofit type.
func (tf *TextFrame) AutofitType() AutofitType {
	if tf.txBody == nil || tf.txBody.BodyPr == nil {
		return AutofitNone
	}
	if tf.txBody.BodyPr.NormAutofit != nil {
		return AutofitNormal
	}
	if tf.txBody.BodyPr.SpAutoFit != nil {
		return AutofitShape
	}
	return AutofitNone
}

// SetAutofitType sets the autofit type.
func (tf *TextFrame) SetAutofitType(t AutofitType) {
	if tf.txBody.BodyPr == nil {
		tf.txBody.BodyPr = &dml.BodyPr{}
	}

	tf.txBody.BodyPr.NoAutofit = nil
	tf.txBody.BodyPr.NormAutofit = nil
	tf.txBody.BodyPr.SpAutoFit = nil

	switch t {
	case AutofitNone:
		tf.txBody.BodyPr.NoAutofit = &dml.NoAutofit{}
	case AutofitNormal:
		tf.txBody.BodyPr.NormAutofit = &dml.NormAutofit{}
	case AutofitShape:
		tf.txBody.BodyPr.SpAutoFit = &dml.SpAutoFit{}
	}
}

// =============================================================================
// TextParagraph
// =============================================================================

// TextParagraph represents a paragraph within a text frame.
type TextParagraph struct {
	textFrame *TextFrame
	p         *dml.P
	index     int
}

// Text returns the combined text of all runs.
func (tp *TextParagraph) Text() string {
	var text string
	for _, r := range tp.p.R {
		text += r.T
	}
	return text
}

// SetText sets the paragraph text, replacing all runs with a single run.
func (tp *TextParagraph) SetText(text string) {
	tp.p.R = []*dml.R{{T: text}}
}

// Runs returns all text runs in the paragraph.
func (tp *TextParagraph) Runs() []*TextRun {
	var result []*TextRun
	for i, r := range tp.p.R {
		result = append(result, &TextRun{
			paragraph: tp,
			r:         r,
			index:     i,
		})
	}
	return result
}

// AddRun adds a new text run to the paragraph.
func (tp *TextParagraph) AddRun() *TextRun {
	r := &dml.R{}
	tp.p.R = append(tp.p.R, r)
	return &TextRun{
		paragraph: tp,
		r:         r,
		index:     len(tp.p.R) - 1,
	}
}

// Level returns the outline level (0-8).
func (tp *TextParagraph) Level() int {
	if tp.p.PPr != nil && tp.p.PPr.Lvl != nil {
		return *tp.p.PPr.Lvl
	}
	return 0
}

// SetLevel sets the outline level (0-8).
func (tp *TextParagraph) SetLevel(level int) {
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	tp.ensurePPr()
	tp.p.PPr.Lvl = &level
}

// BulletType returns the bullet type.
func (tp *TextParagraph) BulletType() BulletType {
	if tp.p.PPr == nil {
		return BulletNone
	}
	if tp.p.PPr.BuAutoNum != nil {
		return BulletAutoNumber
	}
	if tp.p.PPr.BuChar != nil {
		return BulletCharacter
	}
	if tp.p.PPr.BuBlip != nil {
		return BulletPicture
	}
	return BulletNone
}

// SetBulletType sets the bullet type.
func (tp *TextParagraph) SetBulletType(t BulletType) {
	tp.ensurePPr()

	tp.p.PPr.BuNone = nil
	tp.p.PPr.BuAutoNum = nil
	tp.p.PPr.BuChar = nil
	tp.p.PPr.BuBlip = nil

	switch t {
	case BulletNone:
		tp.p.PPr.BuNone = &dml.BuNone{}
	case BulletAutoNumber:
		tp.p.PPr.BuAutoNum = &dml.BuAutoNum{Type: "arabicPeriod"}
	case BulletCharacter:
		tp.p.PPr.BuChar = &dml.BuChar{Char: "•"}
	}
}

// SetBulletCharacter sets a custom bullet character.
func (tp *TextParagraph) SetBulletCharacter(char string) {
	tp.ensurePPr()
	tp.p.PPr.BuNone = nil
	tp.p.PPr.BuAutoNum = nil
	tp.p.PPr.BuBlip = nil
	tp.p.PPr.BuChar = &dml.BuChar{Char: char}
}

// Alignment returns the text alignment.
func (tp *TextParagraph) Alignment() Alignment {
	if tp.p.PPr == nil {
		return AlignmentLeft
	}
	switch tp.p.PPr.Algn {
	case "l":
		return AlignmentLeft
	case "ctr":
		return AlignmentCenter
	case "r":
		return AlignmentRight
	case "just":
		return AlignmentJustify
	default:
		return AlignmentLeft
	}
}

// SetAlignment sets the text alignment.
func (tp *TextParagraph) SetAlignment(a Alignment) {
	tp.ensurePPr()
	switch a {
	case AlignmentLeft:
		tp.p.PPr.Algn = "l"
	case AlignmentCenter:
		tp.p.PPr.Algn = "ctr"
	case AlignmentRight:
		tp.p.PPr.Algn = "r"
	case AlignmentJustify:
		tp.p.PPr.Algn = "just"
	}
}

func (tp *TextParagraph) ensurePPr() {
	if tp.p.PPr == nil {
		tp.p.PPr = &dml.PPr{}
	}
}

// =============================================================================
// TextRun
// =============================================================================

// TextRun represents a run of text with consistent formatting.
type TextRun struct {
	paragraph *TextParagraph
	r         *dml.R
	index     int
}

// Text returns the text of the run.
func (tr *TextRun) Text() string {
	return tr.r.T
}

// SetText sets the text of the run.
func (tr *TextRun) SetText(text string) {
	tr.r.T = text
}

// Bold returns whether the text is bold.
func (tr *TextRun) Bold() bool {
	if tr.r.RPr != nil && tr.r.RPr.B != nil {
		return *tr.r.RPr.B
	}
	return false
}

// SetBold sets whether the text is bold.
func (tr *TextRun) SetBold(b bool) {
	tr.ensureRPr()
	tr.r.RPr.B = &b
}

// Italic returns whether the text is italic.
func (tr *TextRun) Italic() bool {
	if tr.r.RPr != nil && tr.r.RPr.I != nil {
		return *tr.r.RPr.I
	}
	return false
}

// SetItalic sets whether the text is italic.
func (tr *TextRun) SetItalic(i bool) {
	tr.ensureRPr()
	tr.r.RPr.I = &i
}

// Underline returns whether the text is underlined.
func (tr *TextRun) Underline() bool {
	if tr.r.RPr != nil && tr.r.RPr.U != "" && tr.r.RPr.U != "none" {
		return true
	}
	return false
}

// SetUnderline sets whether the text is underlined.
func (tr *TextRun) SetUnderline(u bool) {
	tr.ensureRPr()
	if u {
		tr.r.RPr.U = "sng" // Single underline
	} else {
		tr.r.RPr.U = "none"
	}
}

// FontSize returns the font size in points.
func (tr *TextRun) FontSize() float64 {
	if tr.r.RPr != nil && tr.r.RPr.Sz != nil {
		return float64(*tr.r.RPr.Sz) / 100.0
	}
	return 0
}

// SetFontSize sets the font size in points.
func (tr *TextRun) SetFontSize(points float64) {
	tr.ensureRPr()
	sz := int(points * 100)
	tr.r.RPr.Sz = &sz
}

// FontName returns the font name.
func (tr *TextRun) FontName() string {
	if tr.r.RPr != nil && tr.r.RPr.Latin != nil {
		return tr.r.RPr.Latin.Typeface
	}
	return ""
}

// SetFontName sets the font name.
func (tr *TextRun) SetFontName(name string) {
	tr.ensureRPr()
	tr.r.RPr.Latin = &dml.TextFont{Typeface: name}
	tr.r.RPr.Cs = &dml.TextFont{Typeface: name}
	tr.r.RPr.Ea = &dml.TextFont{Typeface: name}
}

// Color returns the text color as hex (e.g., "FF0000").
func (tr *TextRun) Color() string {
	if tr.r.RPr != nil && tr.r.RPr.SolidFill != nil && tr.r.RPr.SolidFill.SrgbClr != nil {
		return tr.r.RPr.SolidFill.SrgbClr.Val
	}
	return ""
}

// SetColor sets the text color (hex format like "FF0000").
func (tr *TextRun) SetColor(hex string) {
	tr.ensureRPr()
	tr.r.RPr.SolidFill = &dml.SolidFill{
		SrgbClr: &dml.SrgbClr{Val: hex},
	}
}

func (tr *TextRun) ensureRPr() {
	if tr.r.RPr == nil {
		tr.r.RPr = &dml.RPr{}
	}
}
