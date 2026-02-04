// Package document provides style management functionality.
package document

import (
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// StyleType represents the type of style.
type StyleType string

const (
	// StyleTypeParagraph identifies paragraph styles.
	StyleTypeParagraph StyleType = "paragraph"
	// StyleTypeCharacter identifies character styles.
	StyleTypeCharacter StyleType = "character"
	// StyleTypeTable identifies table styles.
	StyleTypeTable StyleType = "table"
	// StyleTypeNumbering identifies numbering styles.
	StyleTypeNumbering StyleType = "numbering"
)

type styles = stylesImpl

// ID returns the style ID.
func (s *styleImpl) ID() string {
	return s.style.StyleID
}

// Name returns the style name.
func (s *styleImpl) Name() string {
	if s.style.Name != nil {
		return s.style.Name.Val
	}
	return ""
}

// SetName sets the style name.
func (s *styleImpl) SetName(name string) {
	if s.style.Name == nil {
		s.style.Name = &wml.StyleName{}
	}
	s.style.Name.Val = name
}

// Type returns the style type.
func (s *styleImpl) Type() StyleType {
	return StyleType(s.style.Type)
}

// BasedOn returns the ID of the style this is based on.
func (s *styleImpl) BasedOn() string {
	if s.style.BasedOn != nil {
		return s.style.BasedOn.Val
	}
	return ""
}

// SetBasedOn sets the parent style ID.
func (s *styleImpl) SetBasedOn(styleID string) {
	if styleID == "" {
		s.style.BasedOn = nil
		return
	}
	s.style.BasedOn = &wml.StyleBasedOn{Val: styleID}
}

// IsDefault returns whether this is a default style.
func (s *styleImpl) IsDefault() bool {
	return s.style.Default != nil && *s.style.Default
}

// SetDefault sets whether this is a default style.
func (s *styleImpl) SetDefault(v bool) {
	s.style.Default = &v
}

// ParagraphProperties returns the paragraph properties for this style.
func (s *styleImpl) ParagraphProperties() *wml.PPr {
	return s.style.PPr
}

// SetParagraphProperties sets the paragraph properties.
func (s *styleImpl) SetParagraphProperties(ppr *wml.PPr) {
	s.style.PPr = ppr
}

// RunProperties returns the run properties for this style.
func (s *styleImpl) RunProperties() *wml.RPr {
	return s.style.RPr
}

// SetRunProperties sets the run properties.
func (s *styleImpl) SetRunProperties(rpr *wml.RPr) {
	s.style.RPr = rpr
}

// =============================================================================
// Document style methods
// =============================================================================

// Styles returns the document styles manager.
func (d *documentImpl) Styles() Styles {
	return &stylesImpl{doc: d}
}

// AddParagraphStyle adds a new paragraph style (legacy helper).
func (d *documentImpl) AddParagraphStyle(id, name string) Style {
	if d == nil {
		return nil
	}
	if style := d.Styles().AddParagraphStyle(id, name); style != nil {
		if s, ok := style.(*styleImpl); ok {
			return s
		}
		return style
	}
	return nil
}

// AddCharacterStyle adds a new character style (legacy helper).
func (d *documentImpl) AddCharacterStyle(id, name string) Style {
	if d == nil {
		return nil
	}
	if style := d.Styles().AddCharacterStyle(id, name); style != nil {
		if s, ok := style.(*styleImpl); ok {
			return s
		}
		return style
	}
	return nil
}

// AddTableStyle adds a new table style (legacy helper).
func (d *documentImpl) AddTableStyle(id, name string) Style {
	if d == nil {
		return nil
	}
	if style := d.Styles().AddTableStyle(id, name); style != nil {
		if s, ok := style.(*styleImpl); ok {
			return s
		}
		return style
	}
	return nil
}

// AddNumberingStyle adds a new numbering style (legacy helper).
func (d *documentImpl) AddNumberingStyle(id, name string) Style {
	if d == nil {
		return nil
	}
	if style := d.Styles().AddNumberingStyle(id, name); style != nil {
		if s, ok := style.(*styleImpl); ok {
			return s
		}
		return style
	}
	return nil
}

// DeleteStyle removes a style by ID (legacy helper).
func (d *documentImpl) DeleteStyle(id string) bool {
	if d == nil {
		return false
	}
	return d.Styles().Delete(id)
}

// DefaultParagraphStyle returns the default paragraph style (legacy helper).
func (d *documentImpl) DefaultParagraphStyle() Style {
	if d == nil {
		return nil
	}
	if style := d.Styles().DefaultParagraphStyle(); style != nil {
		if s, ok := style.(*styleImpl); ok {
			return s
		}
		return style
	}
	return nil
}

// DefaultCharacterStyle returns the default character style (legacy helper).
func (d *documentImpl) DefaultCharacterStyle() Style {
	if d == nil {
		return nil
	}
	if style := d.Styles().DefaultCharacterStyle(); style != nil {
		if s, ok := style.(*styleImpl); ok {
			return s
		}
		return style
	}
	return nil
}

// All returns all styles in the document.
func (s *stylesImpl) All() []Style {
	if s == nil || s.doc == nil || s.doc.styles == nil || s.doc.styles.Styles == nil {
		return nil
	}
	result := make([]Style, len(s.doc.styles.Styles))
	for i, style := range s.doc.styles.Styles {
		result[i] = &styleImpl{doc: s.doc, style: style}
	}
	return result
}

// ByID returns a style by its ID.
func (s *stylesImpl) ByID(id string) Style {
	if s == nil || s.doc == nil || s.doc.styles == nil {
		return nil
	}
	for _, style := range s.doc.styles.Styles {
		if style.StyleID == id {
			return &styleImpl{doc: s.doc, style: style}
		}
	}
	return nil
}

// ByName returns a style by its name.
func (s *stylesImpl) ByName(name string) Style {
	if s == nil || s.doc == nil || s.doc.styles == nil {
		return nil
	}
	for _, style := range s.doc.styles.Styles {
		if style.Name != nil && style.Name.Val == name {
			return &styleImpl{doc: s.doc, style: style}
		}
	}
	return nil
}

// AddParagraphStyle adds a new paragraph style.
func (s *stylesImpl) AddParagraphStyle(id, name string) Style {
	return s.addStyle(id, name, StyleTypeParagraph)
}

// AddCharacterStyle adds a new character style.
func (s *stylesImpl) AddCharacterStyle(id, name string) Style {
	return s.addStyle(id, name, StyleTypeCharacter)
}

// AddTableStyle adds a new table style.
func (s *stylesImpl) AddTableStyle(id, name string) Style {
	return s.addStyle(id, name, StyleTypeTable)
}

// AddNumberingStyle adds a new numbering style.
func (s *stylesImpl) AddNumberingStyle(id, name string) Style {
	return s.addStyle(id, name, StyleTypeNumbering)
}

func (s *stylesImpl) addStyle(id, name string, styleType StyleType) Style {
	if s == nil || s.doc == nil {
		return nil
	}
	if s.doc.styles == nil {
		s.doc.styles = &wml.Styles{}
	}
	style := &wml.Style{
		Type:    string(styleType),
		StyleID: id,
		Name:    &wml.StyleName{Val: name},
	}
	s.doc.styles.Styles = append(s.doc.styles.Styles, style)
	return &styleImpl{doc: s.doc, style: style}
}

// Delete removes a style by ID.
func (s *stylesImpl) Delete(id string) bool {
	if s == nil || s.doc == nil || s.doc.styles == nil {
		return false
	}
	for i, style := range s.doc.styles.Styles {
		if style.StyleID == id {
			s.doc.styles.Styles = append(s.doc.styles.Styles[:i], s.doc.styles.Styles[i+1:]...)
			return true
		}
	}
	return false
}

// DefaultParagraphStyle returns the default paragraph style.
func (s *stylesImpl) DefaultParagraphStyle() Style {
	if s == nil || s.doc == nil || s.doc.styles == nil {
		return nil
	}
	for _, style := range s.doc.styles.Styles {
		if style.Type == "paragraph" && style.Default != nil && *style.Default {
			return &styleImpl{doc: s.doc, style: style}
		}
	}
	return nil
}

// DefaultCharacterStyle returns the default character style.
func (s *stylesImpl) DefaultCharacterStyle() Style {
	if s == nil || s.doc == nil || s.doc.styles == nil {
		return nil
	}
	for _, style := range s.doc.styles.Styles {
		if style.Type == "character" && style.Default != nil && *style.Default {
			return &styleImpl{doc: s.doc, style: style}
		}
	}
	return nil
}

// List returns all styles (alias for All).
func (s *stylesImpl) List() []Style {
	return s.All()
}

// =============================================================================
// Style helper methods
// =============================================================================

// SetBold sets bold formatting on this style.
func (s *styleImpl) SetBold(v bool) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	if v {
		s.style.RPr.B = wml.NewOnOffEnabled()
	} else {
		s.style.RPr.B = nil
	}
}

// SetItalic sets italic formatting on this style.
func (s *styleImpl) SetItalic(v bool) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	if v {
		s.style.RPr.I = wml.NewOnOffEnabled()
	} else {
		s.style.RPr.I = nil
	}
}

// SetFontSize sets the font size in points.
func (s *styleImpl) SetFontSize(points float64) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	halfPoints := int64(points * 2)
	s.style.RPr.Sz = &wml.Sz{Val: halfPoints}
	s.style.RPr.SzCs = &wml.Sz{Val: halfPoints}
}

// SetFontName sets the font name.
func (s *styleImpl) SetFontName(name string) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	s.style.RPr.RFonts = &wml.RFonts{
		Ascii:    name,
		HAnsi:    name,
		EastAsia: name,
		Cs:       name,
	}
}

// SetColor sets the text color.
func (s *styleImpl) SetColor(hex string) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	s.style.RPr.Color = &wml.Color{Val: hex}
}

// SetAlignment sets the paragraph alignment.
func (s *styleImpl) SetAlignment(align string) {
	if s.style.PPr == nil {
		s.style.PPr = &wml.PPr{}
	}
	s.style.PPr.Jc = &wml.Jc{Val: align}
}

// SetSpacingBefore sets the spacing before the paragraph in twips.
func (s *styleImpl) SetSpacingBefore(twips int64) {
	if s.style.PPr == nil {
		s.style.PPr = &wml.PPr{}
	}
	if s.style.PPr.Spacing == nil {
		s.style.PPr.Spacing = &wml.Spacing{}
	}
	s.style.PPr.Spacing.Before = &twips
}

// SetSpacingAfter sets the spacing after the paragraph in twips.
func (s *styleImpl) SetSpacingAfter(twips int64) {
	if s.style.PPr == nil {
		s.style.PPr = &wml.PPr{}
	}
	if s.style.PPr.Spacing == nil {
		s.style.PPr.Spacing = &wml.Spacing{}
	}
	s.style.PPr.Spacing.After = &twips
}

// =============================================================================
// Styles marshaling
// =============================================================================

func (d *documentImpl) saveStyles() error {
	if d.styles == nil {
		return nil
	}
	
	data, err := utils.MarshalXMLWithHeader(d.styles)
	if err != nil {
		return err
	}
	
	_, err = d.pkg.AddPart("word/styles.xml", packaging.ContentTypeStyles, data)
	return err
}
