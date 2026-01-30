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
	StyleTypeParagraph StyleType = "paragraph"
	StyleTypeCharacter StyleType = "character"
	StyleTypeTable     StyleType = "table"
	StyleTypeNumbering StyleType = "numbering"
)

// Style represents a document style.
type Style struct {
	doc   *Document
	style *wml.Style
}

// ID returns the style ID.
func (s *Style) ID() string {
	return s.style.StyleID
}

// Name returns the style name.
func (s *Style) Name() string {
	if s.style.Name != nil {
		return s.style.Name.Val
	}
	return ""
}

// SetName sets the style name.
func (s *Style) SetName(name string) {
	if s.style.Name == nil {
		s.style.Name = &wml.StyleName{}
	}
	s.style.Name.Val = name
}

// Type returns the style type.
func (s *Style) Type() StyleType {
	return StyleType(s.style.Type)
}

// BasedOn returns the ID of the style this is based on.
func (s *Style) BasedOn() string {
	if s.style.BasedOn != nil {
		return s.style.BasedOn.Val
	}
	return ""
}

// SetBasedOn sets the parent style ID.
func (s *Style) SetBasedOn(styleID string) {
	if styleID == "" {
		s.style.BasedOn = nil
		return
	}
	s.style.BasedOn = &wml.StyleBasedOn{Val: styleID}
}

// IsDefault returns whether this is a default style.
func (s *Style) IsDefault() bool {
	return s.style.Default != nil && *s.style.Default
}

// SetDefault sets whether this is a default style.
func (s *Style) SetDefault(v bool) {
	s.style.Default = &v
}

// ParagraphProperties returns the paragraph properties for this style.
func (s *Style) ParagraphProperties() *wml.PPr {
	return s.style.PPr
}

// SetParagraphProperties sets the paragraph properties.
func (s *Style) SetParagraphProperties(ppr *wml.PPr) {
	s.style.PPr = ppr
}

// RunProperties returns the run properties for this style.
func (s *Style) RunProperties() *wml.RPr {
	return s.style.RPr
}

// SetRunProperties sets the run properties.
func (s *Style) SetRunProperties(rpr *wml.RPr) {
	s.style.RPr = rpr
}

// =============================================================================
// Document style methods
// =============================================================================

// Styles returns all styles in the document.
func (d *Document) Styles() []*Style {
	if d.styles == nil || d.styles.Styles == nil {
		return nil
	}
	
	result := make([]*Style, len(d.styles.Styles))
	for i, s := range d.styles.Styles {
		result[i] = &Style{doc: d, style: s}
	}
	return result
}

// StyleByID returns a style by its ID.
func (d *Document) StyleByID(id string) *Style {
	if d.styles == nil {
		return nil
	}
	
	for _, s := range d.styles.Styles {
		if s.StyleID == id {
			return &Style{doc: d, style: s}
		}
	}
	return nil
}

// StyleByName returns a style by its name.
func (d *Document) StyleByName(name string) *Style {
	if d.styles == nil {
		return nil
	}
	
	for _, s := range d.styles.Styles {
		if s.Name != nil && s.Name.Val == name {
			return &Style{doc: d, style: s}
		}
	}
	return nil
}

// AddParagraphStyle adds a new paragraph style.
func (d *Document) AddParagraphStyle(id, name string) *Style {
	return d.addStyle(id, name, StyleTypeParagraph)
}

// AddCharacterStyle adds a new character style.
func (d *Document) AddCharacterStyle(id, name string) *Style {
	return d.addStyle(id, name, StyleTypeCharacter)
}

// AddTableStyle adds a new table style.
func (d *Document) AddTableStyle(id, name string) *Style {
	return d.addStyle(id, name, StyleTypeTable)
}

// AddNumberingStyle adds a new numbering style.
func (d *Document) AddNumberingStyle(id, name string) *Style {
	return d.addStyle(id, name, StyleTypeNumbering)
}

func (d *Document) addStyle(id, name string, styleType StyleType) *Style {
	if d.styles == nil {
		d.styles = &wml.Styles{}
	}
	
	style := &wml.Style{
		Type:    string(styleType),
		StyleID: id,
		Name:    &wml.StyleName{Val: name},
	}
	
	d.styles.Styles = append(d.styles.Styles, style)
	return &Style{doc: d, style: style}
}

// DeleteStyle removes a style by ID.
func (d *Document) DeleteStyle(id string) bool {
	if d.styles == nil {
		return false
	}
	
	for i, s := range d.styles.Styles {
		if s.StyleID == id {
			d.styles.Styles = append(d.styles.Styles[:i], d.styles.Styles[i+1:]...)
			return true
		}
	}
	return false
}

// DefaultParagraphStyle returns the default paragraph style.
func (d *Document) DefaultParagraphStyle() *Style {
	if d.styles == nil {
		return nil
	}
	
	for _, s := range d.styles.Styles {
		if s.Type == "paragraph" && s.Default != nil && *s.Default {
			return &Style{doc: d, style: s}
		}
	}
	return nil
}

// DefaultCharacterStyle returns the default character style.
func (d *Document) DefaultCharacterStyle() *Style {
	if d.styles == nil {
		return nil
	}
	
	for _, s := range d.styles.Styles {
		if s.Type == "character" && s.Default != nil && *s.Default {
			return &Style{doc: d, style: s}
		}
	}
	return nil
}

// =============================================================================
// Style helper methods
// =============================================================================

// SetBold sets bold formatting on this style.
func (s *Style) SetBold(v bool) {
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
func (s *Style) SetItalic(v bool) {
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
func (s *Style) SetFontSize(points float64) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	halfPoints := int64(points * 2)
	s.style.RPr.Sz = &wml.Sz{Val: halfPoints}
	s.style.RPr.SzCs = &wml.Sz{Val: halfPoints}
}

// SetFontName sets the font name.
func (s *Style) SetFontName(name string) {
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
func (s *Style) SetColor(hex string) {
	if s.style.RPr == nil {
		s.style.RPr = &wml.RPr{}
	}
	s.style.RPr.Color = &wml.Color{Val: hex}
}

// SetAlignment sets the paragraph alignment.
func (s *Style) SetAlignment(align string) {
	if s.style.PPr == nil {
		s.style.PPr = &wml.PPr{}
	}
	s.style.PPr.Jc = &wml.Jc{Val: align}
}

// SetSpacingBefore sets the spacing before the paragraph in twips.
func (s *Style) SetSpacingBefore(twips int64) {
	if s.style.PPr == nil {
		s.style.PPr = &wml.PPr{}
	}
	if s.style.PPr.Spacing == nil {
		s.style.PPr.Spacing = &wml.Spacing{}
	}
	s.style.PPr.Spacing.Before = &twips
}

// SetSpacingAfter sets the spacing after the paragraph in twips.
func (s *Style) SetSpacingAfter(twips int64) {
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

func (d *Document) saveStyles() error {
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
