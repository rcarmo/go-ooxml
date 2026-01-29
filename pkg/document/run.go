package document

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Run represents a text run in a paragraph.
type Run struct {
	doc *Document
	r   *wml.R
}

// Text returns the text content of the run.
func (r *Run) Text() string {
	var sb strings.Builder
	for _, elem := range r.r.Content {
		switch v := elem.(type) {
		case *wml.T:
			sb.WriteString(v.Text)
		case *wml.Br:
			sb.WriteString("\n")
		case *wml.Tab:
			sb.WriteString("\t")
		}
	}
	return sb.String()
}

// SetText sets the text content of the run.
func (r *Run) SetText(text string) {
	r.r.Content = []interface{}{wml.NewT(text)}
}

// Bold returns whether the run is bold.
func (r *Run) Bold() bool {
	if r.r.RPr != nil && r.r.RPr.B != nil {
		return r.r.RPr.B.Enabled()
	}
	return false
}

// SetBold sets the bold formatting.
func (r *Run) SetBold(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.B = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.B = nil
	}
}

// Italic returns whether the run is italic.
func (r *Run) Italic() bool {
	if r.r.RPr != nil && r.r.RPr.I != nil {
		return r.r.RPr.I.Enabled()
	}
	return false
}

// SetItalic sets the italic formatting.
func (r *Run) SetItalic(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.I = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.I = nil
	}
}

// Underline returns whether the run is underlined.
func (r *Run) Underline() bool {
	if r.r.RPr != nil && r.r.RPr.U != nil {
		return r.r.RPr.U.Val != "" && r.r.RPr.U.Val != "none"
	}
	return false
}

// SetUnderline sets the underline formatting.
func (r *Run) SetUnderline(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.U = &wml.U{Val: "single"}
	} else {
		r.r.RPr.U = nil
	}
}

// UnderlineStyle returns the underline style.
func (r *Run) UnderlineStyle() string {
	if r.r.RPr != nil && r.r.RPr.U != nil {
		return r.r.RPr.U.Val
	}
	return ""
}

// SetUnderlineStyle sets the underline style (single, double, wave, etc.).
func (r *Run) SetUnderlineStyle(style string) {
	r.ensureRPr()
	if style == "" || style == "none" {
		r.r.RPr.U = nil
	} else {
		r.r.RPr.U = &wml.U{Val: style}
	}
}

// Strike returns whether the run has strikethrough.
func (r *Run) Strike() bool {
	if r.r.RPr != nil && r.r.RPr.Strike != nil {
		return r.r.RPr.Strike.Enabled()
	}
	return false
}

// SetStrike sets the strikethrough formatting.
func (r *Run) SetStrike(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.Strike = wml.NewOnOffEnabled()
	} else {
		r.r.RPr.Strike = nil
	}
}

// FontSize returns the font size in points.
func (r *Run) FontSize() float64 {
	if r.r.RPr != nil && r.r.RPr.Sz != nil {
		return utils.HalfPointsToPoints(r.r.RPr.Sz.Val)
	}
	return 0
}

// SetFontSize sets the font size in points.
func (r *Run) SetFontSize(points float64) {
	r.ensureRPr()
	halfPoints := utils.PointsToHalfPoints(points)
	r.r.RPr.Sz = &wml.Sz{Val: halfPoints}
	r.r.RPr.SzCs = &wml.Sz{Val: halfPoints} // Complex script size
}

// FontName returns the font name.
func (r *Run) FontName() string {
	if r.r.RPr != nil && r.r.RPr.RFonts != nil {
		if r.r.RPr.RFonts.Ascii != "" {
			return r.r.RPr.RFonts.Ascii
		}
	}
	return ""
}

// SetFontName sets the font name.
func (r *Run) SetFontName(name string) {
	r.ensureRPr()
	r.r.RPr.RFonts = &wml.RFonts{
		Ascii:    name,
		HAnsi:    name,
		EastAsia: name,
		Cs:       name,
	}
}

// Color returns the text color as a hex string (without #).
func (r *Run) Color() string {
	if r.r.RPr != nil && r.r.RPr.Color != nil {
		return r.r.RPr.Color.Val
	}
	return ""
}

// SetColor sets the text color (hex string without #).
func (r *Run) SetColor(hex string) {
	r.ensureRPr()
	r.r.RPr.Color = &wml.Color{Val: strings.TrimPrefix(hex, "#")}
}

// Highlight returns the highlight color name.
func (r *Run) Highlight() string {
	if r.r.RPr != nil && r.r.RPr.Highlight != nil {
		return r.r.RPr.Highlight.Val
	}
	return ""
}

// SetHighlight sets the highlight color (yellow, green, cyan, etc.).
func (r *Run) SetHighlight(color string) {
	r.ensureRPr()
	if color == "" {
		r.r.RPr.Highlight = nil
	} else {
		r.r.RPr.Highlight = &wml.Highlight{Val: color}
	}
}

// Style returns the character style ID.
func (r *Run) Style() string {
	if r.r.RPr != nil && r.r.RPr.RStyle != nil {
		return r.r.RPr.RStyle.Val
	}
	return ""
}

// SetStyle sets the character style.
func (r *Run) SetStyle(styleID string) {
	r.ensureRPr()
	if styleID == "" {
		r.r.RPr.RStyle = nil
	} else {
		r.r.RPr.RStyle = &wml.RStyle{Val: styleID}
	}
}

// Superscript returns whether the run is superscript.
func (r *Run) Superscript() bool {
	if r.r.RPr != nil && r.r.RPr.VertAlign != nil {
		return r.r.RPr.VertAlign.Val == "superscript"
	}
	return false
}

// SetSuperscript sets superscript formatting.
func (r *Run) SetSuperscript(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.VertAlign = &wml.VertAlign{Val: "superscript"}
	} else if r.r.RPr.VertAlign != nil && r.r.RPr.VertAlign.Val == "superscript" {
		r.r.RPr.VertAlign = nil
	}
}

// Subscript returns whether the run is subscript.
func (r *Run) Subscript() bool {
	if r.r.RPr != nil && r.r.RPr.VertAlign != nil {
		return r.r.RPr.VertAlign.Val == "subscript"
	}
	return false
}

// SetSubscript sets subscript formatting.
func (r *Run) SetSubscript(v bool) {
	r.ensureRPr()
	if v {
		r.r.RPr.VertAlign = &wml.VertAlign{Val: "subscript"}
	} else if r.r.RPr.VertAlign != nil && r.r.RPr.VertAlign.Val == "subscript" {
		r.r.RPr.VertAlign = nil
	}
}

// AddBreak adds a line break to the run.
func (r *Run) AddBreak() {
	r.r.Content = append(r.r.Content, &wml.Br{})
}

// AddPageBreak adds a page break to the run.
func (r *Run) AddPageBreak() {
	r.r.Content = append(r.r.Content, &wml.Br{Type: "page"})
}

// AddTab adds a tab character to the run.
func (r *Run) AddTab() {
	r.r.Content = append(r.r.Content, &wml.Tab{})
}

func (r *Run) ensureRPr() {
	if r.r.RPr == nil {
		r.r.RPr = &wml.RPr{}
	}
}

// XML returns the underlying WML run for advanced access.
func (r *Run) XML() *wml.R {
	return r.r
}
