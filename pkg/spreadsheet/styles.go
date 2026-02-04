package spreadsheet

import (
	"fmt"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
)

// Styles manages workbook styles.
type stylesImpl struct {
	stylesheet *sml.StyleSheet
	nextNumFmt int
	numFmtMap  map[string]int
}

func newStyles(stylesheet *sml.StyleSheet) *stylesImpl {
	if stylesheet == nil {
		stylesheet = defaultStyleSheet()
	}
	st := &stylesImpl{
		stylesheet: stylesheet,
		numFmtMap:  make(map[string]int),
	}
	st.seedNumFormats()
	return st
}

func (s *stylesImpl) seedNumFormats() {
	if s.stylesheet.NumFmts == nil {
		s.stylesheet.NumFmts = &sml.NumFmts{}
	}
	maxID := 163
	for _, nf := range s.stylesheet.NumFmts.NumFmt {
		s.numFmtMap[nf.FormatCode] = nf.NumFmtID
		if nf.NumFmtID > maxID {
			maxID = nf.NumFmtID
		}
	}
	s.nextNumFmt = maxID + 1
}

// Style returns a new editable cell style derived from the default.
func (s *stylesImpl) Style() CellStyle {
	if s.stylesheet.CellXfs == nil || len(s.stylesheet.CellXfs.Xf) == 0 {
		s.stylesheet.CellXfs = &sml.CellXfs{Xf: []*sml.Xf{{}}}
	}
	return &cellStyleImpl{
		styles: s,
		xf:     &sml.Xf{NumFmtID: 0, FontID: 0, FillID: 0, BorderID: 0, XFID: 0},
	}
}

// ensureFont returns the index for a font definition.
func (s *stylesImpl) ensureFont(font *sml.Font) int {
	if s.stylesheet.Fonts == nil {
		s.stylesheet.Fonts = &sml.Fonts{}
	}
	s.stylesheet.Fonts.Font = append(s.stylesheet.Fonts.Font, font)
	s.stylesheet.Fonts.Count = len(s.stylesheet.Fonts.Font)
	return len(s.stylesheet.Fonts.Font) - 1
}

func (s *stylesImpl) ensureFill(fill *sml.Fill) int {
	if s.stylesheet.Fills == nil {
		s.stylesheet.Fills = &sml.Fills{}
	}
	s.stylesheet.Fills.Fill = append(s.stylesheet.Fills.Fill, fill)
	s.stylesheet.Fills.Count = len(s.stylesheet.Fills.Fill)
	return len(s.stylesheet.Fills.Fill) - 1
}

func (s *stylesImpl) ensureBorder(border *sml.Border) int {
	if s.stylesheet.Borders == nil {
		s.stylesheet.Borders = &sml.Borders{}
	}
	s.stylesheet.Borders.Border = append(s.stylesheet.Borders.Border, border)
	s.stylesheet.Borders.Count = len(s.stylesheet.Borders.Border)
	return len(s.stylesheet.Borders.Border) - 1
}

func (s *stylesImpl) ensureNumFmt(format string) int {
	if format == "" {
		return 0
	}
	if id, ok := s.numFmtMap[format]; ok {
		return id
	}
	id := s.nextNumFmt
	s.nextNumFmt++
	s.stylesheet.NumFmts.NumFmt = append(s.stylesheet.NumFmts.NumFmt, &sml.NumFmt{
		NumFmtID:   id,
		FormatCode: format,
	})
	s.stylesheet.NumFmts.Count = len(s.stylesheet.NumFmts.NumFmt)
	s.numFmtMap[format] = id
	return id
}

func (s *stylesImpl) addCellXf(xf *sml.Xf) int {
	if s.stylesheet.CellXfs == nil {
		s.stylesheet.CellXfs = &sml.CellXfs{}
	}
	s.stylesheet.CellXfs.Xf = append(s.stylesheet.CellXfs.Xf, xf)
	s.stylesheet.CellXfs.Count = len(s.stylesheet.CellXfs.Xf)
	return len(s.stylesheet.CellXfs.Xf) - 1
}

func defaultStyleSheet() *sml.StyleSheet {
	return &sml.StyleSheet{
		Fonts: &sml.Fonts{
			Font: []*sml.Font{{
				Name: &sml.FontName{Val: "Calibri"},
				Family: &sml.FontFamily{Val: 2},
				Sz:   &sml.FontSize{Val: 11},
				Color: &sml.Color{Theme: intPtr(1)},
				Scheme: &sml.FontScheme{Val: "minor"},
			}},
		},
		Fills: &sml.Fills{
			Fill: []*sml.Fill{
				{PatternFill: &sml.PatternFill{}},
				{PatternFill: &sml.PatternFill{PatternType: "gray125"}},
			},
		},
		Borders: &sml.Borders{
			Border: []*sml.Border{{
				Left: &sml.BorderSide{},
				Right: &sml.BorderSide{},
				Top: &sml.BorderSide{},
				Bottom: &sml.BorderSide{},
				Diagonal: &sml.BorderSide{},
			}},
		},
		CellStyleXfs: &sml.CellStyleXfs{Xf: []*sml.Xf{{}}},
		CellXfs:      &sml.CellXfs{Xf: []*sml.Xf{{}}},
		CellStyles:   &sml.CellStyles{CellStyle: []*sml.CellStyle{{Name: "Normal", XFID: 0, BuiltinID: 0}}},
	}
}

func intPtr(v int) *int {
	return &v
}

// CellStyle represents editable style settings.
type cellStyleImpl struct {
	styles *stylesImpl
	xf     *sml.Xf
	font   *sml.Font
	fill   *sml.Fill
	border *sml.Border
}

// FontName returns the font name.
func (cs *cellStyleImpl) FontName() string {
	if cs.font == nil || cs.font.Name == nil {
		return ""
	}
	return cs.font.Name.Val
}

// SetFontName sets the font name.
func (cs *cellStyleImpl) SetFontName(name string) CellStyle {
	cs.ensureFont()
	cs.font.Name = &sml.FontName{Val: name}
	return cs
}

// FontSize returns the font size.
func (cs *cellStyleImpl) FontSize() float64 {
	if cs.font == nil || cs.font.Sz == nil {
		return 0
	}
	return cs.font.Sz.Val
}

// SetFontSize sets the font size.
func (cs *cellStyleImpl) SetFontSize(size float64) CellStyle {
	cs.ensureFont()
	cs.font.Sz = &sml.FontSize{Val: size}
	return cs
}

// Bold returns whether bold is set.
func (cs *cellStyleImpl) Bold() bool {
	return cs.font != nil && cs.font.B != nil
}

// SetBold sets bold.
func (cs *cellStyleImpl) SetBold(v bool) CellStyle {
	cs.ensureFont()
	if v {
		cs.font.B = &sml.BoolVal{Val: "1"}
	} else {
		cs.font.B = nil
	}
	return cs
}

// Italic returns whether italic is set.
func (cs *cellStyleImpl) Italic() bool {
	return cs.font != nil && cs.font.I != nil
}

// SetItalic sets italic.
func (cs *cellStyleImpl) SetItalic(v bool) CellStyle {
	cs.ensureFont()
	if v {
		cs.font.I = &sml.BoolVal{Val: "1"}
	} else {
		cs.font.I = nil
	}
	return cs
}

// FillColor returns the fill color.
func (cs *cellStyleImpl) FillColor() string {
	if cs.fill == nil || cs.fill.PatternFill == nil || cs.fill.PatternFill.FgColor == nil {
		return ""
	}
	return cs.fill.PatternFill.FgColor.RGB
}

// SetFillColor sets the fill color.
func (cs *cellStyleImpl) SetFillColor(hex string) CellStyle {
	cs.ensureFill()
	cs.fill.PatternFill.PatternType = "solid"
	cs.fill.PatternFill.FgColor = &sml.Color{RGB: hex}
	cs.fill.PatternFill.BgColor = &sml.Color{RGB: hex}
	return cs
}

// Border describes a simple border style.
type Border struct {
	Style string
}

// Border returns the current border style.
func (cs *cellStyleImpl) Border() Border {
	if cs.border == nil || cs.border.Left == nil {
		return Border{}
	}
	return Border{Style: cs.border.Left.Style}
}

// SetBorder sets a border style.
func (cs *cellStyleImpl) SetBorder(border Border) CellStyle {
	cs.ensureBorder()
	cs.border.Left.Style = border.Style
	cs.border.Right.Style = border.Style
	cs.border.Top.Style = border.Style
	cs.border.Bottom.Style = border.Style
	return cs
}

// Alignment represents alignment options for cell styles.
type Alignment string

// HorizontalAlignment returns the horizontal alignment.
func (cs *cellStyleImpl) HorizontalAlignment() Alignment {
	if cs.xf == nil || cs.xf.Alignment == nil {
		return ""
	}
	return Alignment(cs.xf.Alignment.Horizontal)
}

// SetHorizontalAlignment sets the horizontal alignment.
func (cs *cellStyleImpl) SetHorizontalAlignment(a Alignment) CellStyle {
	if cs.xf.Alignment == nil {
		cs.xf.Alignment = &sml.Alignment{}
	}
	cs.xf.Alignment.Horizontal = string(a)
	cs.xf.ApplyAlignment = boolPtr(true)
	return cs
}

// VerticalAlignment returns the vertical alignment.
func (cs *cellStyleImpl) VerticalAlignment() Alignment {
	if cs.xf == nil || cs.xf.Alignment == nil {
		return ""
	}
	return Alignment(cs.xf.Alignment.Vertical)
}

// SetVerticalAlignment sets the vertical alignment.
func (cs *cellStyleImpl) SetVerticalAlignment(a Alignment) CellStyle {
	if cs.xf.Alignment == nil {
		cs.xf.Alignment = &sml.Alignment{}
	}
	cs.xf.Alignment.Vertical = string(a)
	cs.xf.ApplyAlignment = boolPtr(true)
	return cs
}

// NumberFormat returns the number format string.
func (cs *cellStyleImpl) NumberFormat() string {
	if cs.xf == nil {
		return ""
	}
	return cs.styles.formatCode(cs.xf.NumFmtID)
}

// SetNumberFormat sets the number format string.
func (cs *cellStyleImpl) SetNumberFormat(format string) CellStyle {
	if cs.xf == nil {
		return cs
	}
	cs.xf.NumFmtID = cs.styles.ensureNumFmt(format)
	cs.xf.ApplyNumberFormat = boolPtr(true)
	return cs
}

func (cs *cellStyleImpl) ensureFont() {
	if cs.font == nil {
		cs.font = &sml.Font{}
	}
}

func (cs *cellStyleImpl) ensureFill() {
	if cs.fill == nil {
		cs.fill = &sml.Fill{PatternFill: &sml.PatternFill{}}
	}
}

func (cs *cellStyleImpl) ensureBorder() {
	if cs.border == nil {
		cs.border = &sml.Border{
			Left:   &sml.BorderSide{},
			Right:  &sml.BorderSide{},
			Top:    &sml.BorderSide{},
			Bottom: &sml.BorderSide{},
		}
	}
}

func (cs *cellStyleImpl) finalize() int {
	if cs.styles == nil {
		return 0
	}
	if cs.xf.XFID == 0 {
		cs.xf.XFID = 0
	}
	if cs.font != nil {
		cs.xf.FontID = cs.styles.ensureFont(cs.font)
	}
	if cs.fill != nil {
		cs.xf.FillID = cs.styles.ensureFill(cs.fill)
	}
	if cs.border != nil {
		cs.xf.BorderID = cs.styles.ensureBorder(cs.border)
	}
	return cs.styles.addCellXf(cs.xf)
}

func (s *stylesImpl) formatCode(id int) string {
	if s.stylesheet.NumFmts == nil {
		return ""
	}
	for _, nf := range s.stylesheet.NumFmts.NumFmt {
		if nf.NumFmtID == id {
			return nf.FormatCode
		}
	}
	if id == 0 {
		return ""
	}
	return fmt.Sprintf("%d", id)
}

func boolPtr(v bool) *bool {
	return &v
}
