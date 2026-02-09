package document

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Text returns the text content of the run.
func (r *runImpl) Text() string {
	var sb strings.Builder
	for _, elem := range r.r.Content {
		switch v := elem.(type) {
		case *wml.T:
			sb.WriteString(v.Text)
		case *wml.Br:
			sb.WriteString("\n")
		case *wml.Tab:
			sb.WriteString("\t")
		case *wml.Sym:
			sb.WriteRune(symToRune(v.Char))
		}
	}
	return sb.String()
}

// SetText sets the text content of the run.
func (r *runImpl) SetText(text string) {
	r.r.Content = []interface{}{wml.NewT(text)}
}

// Properties returns the run properties.
func (r *runImpl) Properties() RunProperties {
	r.ensureRPr()
	if r.r.RPr == nil {
		return RunProperties{}
	}
	return *r.r.RPr
}

// Bold returns whether the run is bold.
func (r *runImpl) Bold() bool {
	if r.r.RPr != nil && r.r.RPr.B != nil {
		return r.r.RPr.B.Enabled()
	}
	return false
}

// SetBold sets the bold formatting.
func (r *runImpl) SetBold(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.B = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.B = nil
	}
}

// Italic returns whether the run is italic.
func (r *runImpl) Italic() bool {
	if r.r.RPr != nil && r.r.RPr.I != nil {
		return r.r.RPr.I.Enabled()
	}
	return false
}

// SetItalic sets the italic formatting.
func (r *runImpl) SetItalic(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.I = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.I = nil
	}
}

// Underline returns whether the run is underlined.
func (r *runImpl) Underline() bool {
	if r.r.RPr != nil && r.r.RPr.U != nil {
		return r.r.RPr.U.Val != "" && r.r.RPr.U.Val != "none"
	}
	return false
}

// SetUnderline sets the underline formatting.
func (r *runImpl) SetUnderline(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.U = &wml.U{Val: "single"}
	} else {
		r.r.RPr.U = nil
	}
}

// UnderlineStyle returns the underline style.
func (r *runImpl) UnderlineStyle() string {
	if r.r.RPr != nil && r.r.RPr.U != nil {
		return r.r.RPr.U.Val
	}
	return ""
}

// SetUnderlineStyle sets the underline style (single, double, wave, etc.).
func (r *runImpl) SetUnderlineStyle(style string) {
	r.ensureRPr()
	if style == "" || style == "none" {
		r.r.RPr.U = nil
	} else {
		r.r.RPr.U = &wml.U{Val: style}
	}
}

// Strike returns whether the run has strikethrough.
func (r *runImpl) Strike() bool {
	if r.r.RPr != nil && r.r.RPr.Strike != nil {
		return r.r.RPr.Strike.Enabled()
	}
	return false
}

// SetStrike sets the strikethrough formatting.
func (r *runImpl) SetStrike(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Strike = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Strike = nil
	}
}

// DoubleStrike returns whether the run has double strikethrough.
func (r *runImpl) DoubleStrike() bool {
	if r.r.RPr != nil && r.r.RPr.Dstrike != nil {
		return r.r.RPr.Dstrike.Enabled()
	}
	return false
}

// SetDoubleStrike sets the double strikethrough formatting.
func (r *runImpl) SetDoubleStrike(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Dstrike = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Dstrike = nil
	}
}

// Caps returns whether the run uses all caps.
func (r *runImpl) Caps() bool {
	if r.r.RPr != nil && r.r.RPr.Caps != nil {
		return r.r.RPr.Caps.Enabled()
	}
	return false
}

// SetCaps sets all caps formatting.
func (r *runImpl) SetCaps(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Caps = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Caps = nil
	}
}

// SmallCaps returns whether the run uses small caps.
func (r *runImpl) SmallCaps() bool {
	if r.r.RPr != nil && r.r.RPr.SmallCaps != nil {
		return r.r.RPr.SmallCaps.Enabled()
	}
	return false
}

// SetSmallCaps sets small caps formatting.
func (r *runImpl) SetSmallCaps(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.SmallCaps = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.SmallCaps = nil
	}
}

// Outline returns whether the run uses outline text.
func (r *runImpl) Outline() bool {
	if r.r.RPr != nil && r.r.RPr.Outline != nil {
		return r.r.RPr.Outline.Enabled()
	}
	return false
}

// SetOutline sets outline text formatting.
func (r *runImpl) SetOutline(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Outline = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Outline = nil
	}
}

// Shadow returns whether the run uses shadow formatting.
func (r *runImpl) Shadow() bool {
	if r.r.RPr != nil && r.r.RPr.Shadow != nil {
		return r.r.RPr.Shadow.Enabled()
	}
	return false
}

// SetShadow sets shadow formatting.
func (r *runImpl) SetShadow(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Shadow = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Shadow = nil
	}
}

// Emboss returns whether the run uses emboss formatting.
func (r *runImpl) Emboss() bool {
	if r.r.RPr != nil && r.r.RPr.Emboss != nil {
		return r.r.RPr.Emboss.Enabled()
	}
	return false
}

// SetEmboss sets emboss formatting.
func (r *runImpl) SetEmboss(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Emboss = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Emboss = nil
	}
}

// Imprint returns whether the run uses imprint formatting.
func (r *runImpl) Imprint() bool {
	if r.r.RPr != nil && r.r.RPr.Imprint != nil {
		return r.r.RPr.Imprint.Enabled()
	}
	return false
}

// SetImprint sets imprint formatting.
func (r *runImpl) SetImprint(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Imprint = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Imprint = nil
	}
}

// Vanish returns whether the run uses hidden text formatting.
func (r *runImpl) Vanish() bool {
	if r.r.RPr != nil && r.r.RPr.Vanish != nil {
		return r.r.RPr.Vanish.Enabled()
	}
	return false
}

// SetVanish sets hidden text formatting.
func (r *runImpl) SetVanish(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Vanish = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Vanish = nil
	}
}

// FontSize returns the font size in points.
func (r *runImpl) FontSize() float64 {
	if r.r.RPr != nil && r.r.RPr.Sz != nil {
		return utils.HalfPointsToPoints(r.r.RPr.Sz.Val)
	}
	return 0
}

// SetFontSize sets the font size in points.
func (r *runImpl) SetFontSize(points float64) {
	r.ensureRPr()
	halfPoints := utils.PointsToHalfPoints(points)
	r.r.RPr.Sz = &wml.Sz{Val: halfPoints}
	r.r.RPr.SzCs = &wml.Sz{Val: halfPoints} // Complex script size
}

// FontName returns the font name.
func (r *runImpl) FontName() string {
	if r.r.RPr != nil && r.r.RPr.RFonts != nil {
		if r.r.RPr.RFonts.Ascii != "" {
			return r.r.RPr.RFonts.Ascii
		}
	}
	return ""
}

// SetFontName sets the font name.
func (r *runImpl) SetFontName(name string) {
	r.ensureRPr()
	r.r.RPr.RFonts = &wml.RFonts{
		Ascii:    name,
		HAnsi:    name,
		EastAsia: name,
		Cs:       name,
	}
}

// Color returns the text color as a hex string (without #).
func (r *runImpl) Color() string {
	if r.r.RPr != nil && r.r.RPr.Color != nil {
		return r.r.RPr.Color.Val
	}
	return ""
}

// SetColor sets the text color (hex string without #).
func (r *runImpl) SetColor(hex string) {
	r.ensureRPr()
	r.r.RPr.Color = &wml.Color{Val: strings.TrimPrefix(hex, "#")}
}

// Highlight returns the highlight color name.
func (r *runImpl) Highlight() string {
	if r.r.RPr != nil && r.r.RPr.Highlight != nil {
		return r.r.RPr.Highlight.Val
	}
	return ""
}

// SetHighlight sets the highlight color (yellow, green, cyan, etc.).
func (r *runImpl) SetHighlight(color string) {
	r.ensureRPr()
	if color == "" {
		r.r.RPr.Highlight = nil
	} else {
		r.r.RPr.Highlight = &wml.Highlight{Val: color}
	}
}

// Style returns the character style ID.
func (r *runImpl) Style() string {
	if r.r.RPr != nil && r.r.RPr.RStyle != nil {
		return r.r.RPr.RStyle.Val
	}
	return ""
}

// SetStyle sets the character style.
func (r *runImpl) SetStyle(styleID string) {
	r.ensureRPr()
	if styleID == "" {
		r.r.RPr.RStyle = nil
	} else {
		r.r.RPr.RStyle = &wml.RStyle{Val: styleID}
	}
}

// Superscript returns whether the run is superscript.
func (r *runImpl) Superscript() bool {
	if r.r.RPr != nil && r.r.RPr.VertAlign != nil {
		return r.r.RPr.VertAlign.Val == "superscript"
	}
	return false
}

// SetSuperscript sets superscript formatting.
func (r *runImpl) SetSuperscript(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.VertAlign = &wml.VertAlign{Val: "superscript"}
	} else if r.r.RPr.VertAlign != nil && r.r.RPr.VertAlign.Val == "superscript" {
		r.r.RPr.VertAlign = nil
	}
}

// Subscript returns whether the run is subscript.
func (r *runImpl) Subscript() bool {
	if r.r.RPr != nil && r.r.RPr.VertAlign != nil {
		return r.r.RPr.VertAlign.Val == "subscript"
	}
	return false
}

// SetSubscript sets subscript formatting.
func (r *runImpl) SetSubscript(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.VertAlign = &wml.VertAlign{Val: "subscript"}
	} else if r.r.RPr.VertAlign != nil && r.r.RPr.VertAlign.Val == "subscript" {
		r.r.RPr.VertAlign = nil
	}
}

// AddBreak adds a line break to the run.
func (r *runImpl) AddBreak() {
	r.r.Content = append(r.r.Content, &wml.Br{})
}

// AddPageBreak adds a page break to the run.
func (r *runImpl) AddPageBreak() {
	r.r.Content = append(r.r.Content, &wml.Br{Type: "page"})
}

// AddTab adds a tab character to the run.
func (r *runImpl) AddTab() {
	r.r.Content = append(r.r.Content, &wml.Tab{})
}

// AddSymbol adds a symbol run element.
func (r *runImpl) AddSymbol(font, char string) {
	if char == "" {
		return
	}
	r.r.Content = append(r.r.Content, &wml.Sym{Font: font, Char: char})
}

// AddLastRenderedPageBreak adds a last rendered page break marker.
func (r *runImpl) AddLastRenderedPageBreak() {
	r.r.Content = append(r.r.Content, &wml.LastRenderedPageBreak{})
}

func (r *runImpl) ensureRPr() {
	if r.r.RPr == nil {
		r.r.RPr = &wml.RPr{}
	}
}

// XML returns the underlying WML run for advanced access.
func (r *runImpl) XML() *wml.R {
	return r.r
}
