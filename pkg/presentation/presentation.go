// Package presentation provides a high-level API for working with PowerPoint presentations.
package presentation

import (
	"fmt"
	"io"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// Standard slide sizes in EMUs (English Metric Units).
const (
	// Standard 4:3 slide dimensions
	SlideWidth4x3  int64 = 9144000  // 10 inches
	SlideHeight4x3 int64 = 6858000  // 7.5 inches

	// Widescreen 16:9 slide dimensions
	SlideWidth16x9  int64 = 12192000 // 13.333 inches
	SlideHeight16x9 int64 = 6858000  // 7.5 inches

	// Widescreen 16:10 slide dimensions
	SlideWidth16x10  int64 = 10972800 // 12 inches
	SlideHeight16x10 int64 = 6858000  // 7.5 inches
)

// Presentation represents a PowerPoint presentation.
type presentationImpl struct {
	pkg          *packaging.Package
	presentation *pml.Presentation
	slides       []*slideImpl
	path         string
	nextSlideID  int
}

// New creates a new empty presentation with standard 4:3 dimensions.
func New() (Presentation, error) {
	return NewWithSize(SlideWidth4x3, SlideHeight4x3)
}

// NewWidescreen creates a new presentation with 16:9 widescreen dimensions.
func NewWidescreen() (Presentation, error) {
	return NewWithSize(SlideWidth16x9, SlideHeight16x9)
}

// NewWithSize creates a new presentation with specified dimensions in EMUs.
func NewWithSize(width, height int64) (Presentation, error) {
	p := &presentationImpl{
		pkg: packaging.New(),
		presentation: &pml.Presentation{
			XMLNS_R: pml.NSR,
			SldSz: &pml.SldSz{
				Cx:   width,
				Cy:   height,
				Type: getSizeType(width, height),
			},
			NotesSz: &pml.NotesSz{
				Cx: 6858000,  // Standard notes width
				Cy: 9144000,  // Standard notes height
			},
			SldIdLst: &pml.SldIdLst{},
		},
		slides:      make([]*slideImpl, 0),
		nextSlideID: 256, // PowerPoint typically starts slide IDs at 256
	}

	if err := p.initPackage(); err != nil {
		return nil, err
	}

	return p, nil
}

// Open opens an existing presentation from a file path.
func Open(path string) (Presentation, error) {
	pkg, err := packaging.Open(path)
	if err != nil {
		return nil, err
	}

	p, err := openFromPackage(pkg)
	if err != nil {
		return nil, err
	}
	p.path = path
	return p, nil
}

// OpenReader opens a presentation from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (Presentation, error) {
	pkg, err := packaging.OpenReader(r, size)
	if err != nil {
		return nil, err
	}
	return openFromPackage(pkg)
}

func openFromPackage(pkg *packaging.Package) (*presentationImpl, error) {
	p := &presentationImpl{
		pkg:         pkg,
		slides:      make([]*slideImpl, 0),
		nextSlideID: 256,
	}

	// Parse presentation.xml
	if err := p.parsePresentation(); err != nil {
		return nil, err
	}

	// Parse slides
	if err := p.parseSlides(); err != nil {
		return nil, err
	}

	return p, nil
}

// Save saves the presentation to its original path.
func (p *presentationImpl) Save() error {
	if p.path == "" {
		return fmt.Errorf("no path set, use SaveAs")
	}
	return p.SaveAs(p.path)
}

// SaveAs saves the presentation to a new path.
func (p *presentationImpl) SaveAs(path string) error {
	if err := p.updatePackage(); err != nil {
		return err
	}
	return p.pkg.SaveAs(path)
}

// Close closes the presentation and releases resources.
func (p *presentationImpl) Close() error {
	return p.pkg.Close()
}

// CoreProperties returns the presentation core properties.
func (p *presentationImpl) CoreProperties() (*common.CoreProperties, error) {
	return p.pkg.CoreProperties()
}

// SetCoreProperties sets the presentation core properties.
func (p *presentationImpl) SetCoreProperties(props *common.CoreProperties) error {
	return p.pkg.SetCoreProperties(props)
}

// Properties returns the presentation properties.
func (p *presentationImpl) Properties() PresentationProperties {
	props, _ := p.pkg.CoreProperties()
	if props == nil {
		return PresentationProperties{}
	}
	return *props
}

// Masters returns slide masters (not implemented).
func (p *presentationImpl) Masters() []SlideMaster {
	return nil
}

// Layouts returns slide layouts (not implemented).
func (p *presentationImpl) Layouts() []SlideLayout {
	return nil
}

// =============================================================================
// Slide access
// =============================================================================

// Slides returns all slides in the presentation.
func (p *presentationImpl) Slides() []Slide {
	slides := make([]Slide, len(p.slides))
	for i, slide := range p.slides {
		slides[i] = slide
	}
	return slides
}

// SlidesRaw returns all slides in the presentation as concrete types.
func (p *presentationImpl) SlidesRaw() []*slideImpl {
	return p.slides
}

// Slide returns a slide by index (1-based).
func (p *presentationImpl) Slide(index int) (Slide, error) {
	index--
	if index < 0 || index >= len(p.slides) {
		return nil, fmt.Errorf("slide index %d out of range", index+1)
	}
	return p.slides[index], nil
}

// SlideCount returns the number of slides.
func (p *presentationImpl) SlideCount() int {
	return len(p.slides)
}

// AddSlide adds a new blank slide at the end.
func (p *presentationImpl) AddSlide(layoutIndex int) Slide {
	return p.InsertSlide(len(p.slides)+1, layoutIndex)
}

// InsertSlide inserts a new blank slide at the specified index (0-based).
func (p *presentationImpl) InsertSlide(index, layoutIndex int) Slide {
	index--
	if index < 0 {
		index = 0
	}
	if index > len(p.slides) {
		index = len(p.slides)
	}

	slideNum := len(p.slides) + 1
	relID := fmt.Sprintf("rId%d", slideNum+10) // Start relationship IDs at 11

	slide := &slideImpl{
		pres:  p,
		slide: createBlankSlide(),
		id:    p.nextSlideID,
		relID: relID,
		index: index,
	}
	p.nextSlideID++

	// Insert into slide list
	if index >= len(p.slides) {
		p.slides = append(p.slides, slide)
	} else {
		p.slides = append(p.slides[:index+1], p.slides[index:]...)
		p.slides[index] = slide
	}

	// Update indices for all slides after insertion point
	for i := index; i < len(p.slides); i++ {
		p.slides[i].index = i
	}

	// Add to presentation slide list
	if p.presentation.SldIdLst == nil {
		p.presentation.SldIdLst = &pml.SldIdLst{}
	}
	newSldID := &pml.SldId{
		ID:  slide.id,
		RID: relID,
	}
	if index >= len(p.presentation.SldIdLst.SldId) {
		p.presentation.SldIdLst.SldId = append(p.presentation.SldIdLst.SldId, newSldID)
	} else {
		p.presentation.SldIdLst.SldId = append(p.presentation.SldIdLst.SldId[:index+1], p.presentation.SldIdLst.SldId[index:]...)
		p.presentation.SldIdLst.SldId[index] = newSldID
	}

	return slide
}

// DeleteSlide removes a slide at the specified index (1-based).
func (p *presentationImpl) DeleteSlide(index int) error {
	index--
	if index < 0 || index >= len(p.slides) {
		return fmt.Errorf("slide index %d out of range", index+1)
	}

	// Remove from slides slice
	p.slides = append(p.slides[:index], p.slides[index+1:]...)

	// Update indices
	for i := index; i < len(p.slides); i++ {
		p.slides[i].index = i
	}

	// Remove from presentation slide list
	if p.presentation.SldIdLst != nil && index < len(p.presentation.SldIdLst.SldId) {
		p.presentation.SldIdLst.SldId = append(
			p.presentation.SldIdLst.SldId[:index],
			p.presentation.SldIdLst.SldId[index+1:]...,
		)
	}

	return nil
}

// DuplicateSlide creates a copy of the slide at the specified index (1-based).
func (p *presentationImpl) DuplicateSlide(index int) Slide {
	index--
	if index < 0 || index >= len(p.slides) {
		return nil
	}

	source := p.slides[index]
	newSlide := p.InsertSlide(index+2, 0)

	// Copy content (deep copy of the slide structure)
	if source.slide.CSld != nil {
		if s, ok := newSlide.(*slideImpl); ok {
			s.slide.CSld = copyCSld(source.slide.CSld)
		}
	}

	return newSlide
}

// ReorderSlides reorders slides according to the new order.
// newOrder contains the new indices, e.g., [2, 0, 1] moves slide 2 to position 0.
func (p *presentationImpl) ReorderSlides(newOrder []int) error {
	if len(newOrder) != len(p.slides) {
		return fmt.Errorf("newOrder length %d doesn't match slide count %d", len(newOrder), len(p.slides))
	}

	// Validate indices (1-based)
	used := make(map[int]bool)
	for _, idx := range newOrder {
		if idx < 1 || idx > len(p.slides) {
			return fmt.Errorf("invalid index %d in newOrder", idx)
		}
		if used[idx] {
			return fmt.Errorf("duplicate index %d in newOrder", idx)
		}
		used[idx] = true
	}

	// Create new slides slice
	newSlides := make([]*slideImpl, len(p.slides))
	newSldIds := make([]*pml.SldId, len(p.slides))
	for newIdx, oldIdx := range newOrder {
		oldIdx--
		newSlides[newIdx] = p.slides[oldIdx]
		newSlides[newIdx].index = newIdx
		if p.presentation.SldIdLst != nil && oldIdx < len(p.presentation.SldIdLst.SldId) {
			newSldIds[newIdx] = p.presentation.SldIdLst.SldId[oldIdx]
		}
	}

	p.slides = newSlides
	if p.presentation.SldIdLst != nil {
		p.presentation.SldIdLst.SldId = newSldIds
	}

	return nil
}

// =============================================================================
// Properties
// =============================================================================

// SlideSize returns the slide dimensions in EMUs.
func (p *presentationImpl) SlideSize() (width, height int64) {
	if p.presentation.SldSz != nil {
		return p.presentation.SldSz.Cx, p.presentation.SldSz.Cy
	}
	return SlideWidth4x3, SlideHeight4x3
}

// SetSlideSize sets the slide dimensions in EMUs.
func (p *presentationImpl) SetSlideSize(width, height int64) error {
	if p.presentation.SldSz == nil {
		p.presentation.SldSz = &pml.SldSz{}
	}
	p.presentation.SldSz.Cx = width
	p.presentation.SldSz.Cy = height
	p.presentation.SldSz.Type = getSizeType(width, height)
	return nil
}

// =============================================================================
// Internal methods
// =============================================================================

func (p *presentationImpl) initPackage() error {
	// Content types are added automatically when parts are added
	// Add main relationship
	p.pkg.AddRelationship("/", "ppt/presentation.xml", packaging.RelTypeOfficeDocument)

	return nil
}

func (p *presentationImpl) parsePresentation() error {
	part, err := p.pkg.GetPart(packaging.PresentationPath)
	if err != nil {
		return err
	}
	data, err := part.Content()
	if err != nil {
		return err
	}

	p.presentation = &pml.Presentation{}
	return utils.UnmarshalXML(data, p.presentation)
}

func (p *presentationImpl) parseSlides() error {
	if p.presentation.SldIdLst == nil {
		return nil
	}

	// Get relationships to find slide parts
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	if rels == nil {
		return nil
	}

	for i, sldId := range p.presentation.SldIdLst.SldId {
		// Find the relationship
		var slidePath string
		for _, rel := range rels.Relationships {
			if rel.ID == sldId.RID {
				slidePath = "ppt/" + rel.Target
				break
			}
		}

		if slidePath == "" {
			continue
		}

		// Parse slide
		part, err := p.pkg.GetPart(slidePath)
		if err != nil {
			continue
		}
		data, err := part.Content()
		if err != nil {
			continue
		}

		slide := &pml.Sld{}
		if err := utils.UnmarshalXML(data, slide); err != nil {
			continue
		}

		p.slides = append(p.slides, &slideImpl{
			pres:  p,
			slide: slide,
			id:    sldId.ID,
			relID: sldId.RID,
			index: i,
			path:  slidePath,
		})

		if sldId.ID >= p.nextSlideID {
			p.nextSlideID = sldId.ID + 1
		}
	}

	return nil
}

func (p *presentationImpl) updatePackage() error {
	// Save presentation.xml
	data, err := utils.MarshalXMLWithHeader(p.presentation)
	if err != nil {
		return err
	}
	if _, err := p.pkg.AddPart(packaging.PresentationPath, packaging.ContentTypePresentation, data); err != nil {
		return err
	}

	// Save each slide
	for i, slide := range p.slides {
		slidePath := fmt.Sprintf("ppt/slides/slide%d.xml", i+1)
		slideData, err := utils.MarshalXMLWithHeader(slide.slide)
		if err != nil {
			return err
		}
		if _, err := p.pkg.AddPart(slidePath, packaging.ContentTypeSlide, slideData); err != nil {
			return err
		}

		// Add relationship with the same ID used in presentation.xml
		rels := p.pkg.GetRelationships(packaging.PresentationPath)
		rels.AddWithID(slide.relID, packaging.RelTypeSlide, "slides/slide"+fmt.Sprintf("%d.xml", i+1), packaging.TargetModeInternal)

		// Save notes if present
		if slide.notes != nil {
			notesPath := fmt.Sprintf("ppt/notesSlides/notesSlide%d.xml", i+1)
			notesData, err := utils.MarshalXMLWithHeader(slide.notes)
			if err != nil {
				return err
			}
			if _, err := p.pkg.AddPart(notesPath, packaging.ContentTypeNotesSlide, notesData); err != nil {
				return err
			}
		}
	}

	return nil
}

func getSizeType(width, height int64) string {
	switch {
	case width == SlideWidth4x3 && height == SlideHeight4x3:
		return pml.SldSzScreen4x3
	case width == SlideWidth16x9 && height == SlideHeight16x9:
		return pml.SldSzScreen16x9
	case width == SlideWidth16x10 && height == SlideHeight16x10:
		return pml.SldSzScreen16x10
	default:
		return pml.SldSzCustom
	}
}

func createBlankSlide() *pml.Sld {
	return &pml.Sld{
		CSld: &pml.CSld{
			SpTree: &pml.SpTree{
				NvGrpSpPr: &pml.NvGrpSpPr{
					CNvPr: &pml.CNvPr{ID: 1, Name: ""},
					CNvGrpSpPr: &pml.CNvGrpSpPr{},
					NvPr:       &pml.NvPr{},
				},
				GrpSpPr: &pml.GrpSpPr{},
			},
		},
	}
}

func copyCSld(src *pml.CSld) *pml.CSld {
	if src == nil {
		return nil
	}
	data, err := utils.MarshalXMLWithHeader(src)
	if err != nil {
		return &pml.CSld{Name: src.Name}
	}
	copied := &pml.CSld{}
	if err := utils.UnmarshalXML(data, copied); err != nil {
		return &pml.CSld{Name: src.Name}
	}
	return copied
}
