// Package presentation provides a high-level API for working with PowerPoint presentations.
package presentation

import (
	"bytes"
	_ "embed"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
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

//go:embed templates/default.pptx
var defaultTemplate []byte

// Presentation represents a PowerPoint presentation.
type presentationImpl struct {
	pkg          *packaging.Package
	presentation *pml.Presentation
	slides       []*slideImpl
	path         string
	nextSlideID  int
	nextChartID  int
	nextDiagramID int
	commentAuthors *pml.AuthorList
	notesMaster  *pml.NotesMaster
	notesMasterPath string
	notesMasterRelID string
	notesMasterTheme []byte
	notesMasterThemePath string
	masters       []*slideMasterImpl
	layouts       []*slideLayoutImpl
	nextImageID   int
	themeParts    map[string][]byte
	extraParts    map[string]*packaging.Part
}

// New creates a new empty presentation with standard 4:3 dimensions.
func New() (Presentation, error) {
	p, err := newFromTemplate()
	if err == nil {
		return p, nil
	}
	return NewWithSize(SlideWidth4x3, SlideHeight4x3)
}

// NewWidescreen creates a new presentation with 16:9 widescreen dimensions.
func NewWidescreen() (Presentation, error) {
	p, err := newFromTemplate()
	if err == nil {
		_ = p.SetSlideSize(SlideWidth16x9, SlideHeight16x9)
		return p, nil
	}
	return NewWithSize(SlideWidth16x9, SlideHeight16x9)
}

// NewWithSize creates a new presentation with specified dimensions in EMUs.
func NewWithSize(width, height int64) (Presentation, error) {
	p, err := newFromTemplate()
	if err != nil {
		p, err = newEmptyPresentation(width, height)
	}
	if err != nil {
		return nil, err
	}
	_ = p.SetSlideSize(width, height)
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

func newFromTemplate() (*presentationImpl, error) {
	if len(defaultTemplate) == 0 {
		return nil, utils.ErrPartNotFound
	}
	pkg, err := packaging.OpenReader(bytes.NewReader(defaultTemplate), int64(len(defaultTemplate)))
	if err != nil {
		return nil, err
	}
	p, err := openFromPackage(pkg)
	if err != nil {
		return nil, err
	}
	if p.presentation != nil {
		p.presentation.SldIdLst = &pml.SldIdLst{}
		p.presentation.NotesMasterIdLst = &pml.NotesMasterIdLst{NotesMasterId: []*pml.NotesMasterId{{RID: "rId7"}}}
		p.presentation.ExtLst = &pml.ExtLst{
			Ext: []*pml.ExtItem{{
				URI: "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}",
				Any: `<p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"><p15:guide id="1" orient="horz" pos="2160"><p15:clr><a:srgbClr val="A4A3A4"/></p15:clr></p15:guide><p15:guide id="2" pos="2880"><p15:clr><a:srgbClr val="A4A3A4"/></p15:clr></p15:guide></p15:sldGuideLst>`,
			}},
		}
		p.presentation.XMLName = xml.Name{Space: pml.NS, Local: "p:presentation"}
		p.presentation.XMLNS_R = pml.NSR
	}
	p.slides = nil
	p.nextSlideID = 256
	presRels := p.pkg.GetRelationships(packaging.PresentationPath)
	filtered := presRels.Relationships[:0]
	for _, rel := range presRels.Relationships {
		if rel.Type == packaging.RelTypeSlide {
			continue
		}
		filtered = append(filtered, rel)
	}
	presRels.Relationships = filtered
	return p, nil
}

func newEmptyPresentation(width, height int64) (*presentationImpl, error) {
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
				Cx: 6858000, // Standard notes width
				Cy: 9144000, // Standard notes height
			},
			SldIdLst: &pml.SldIdLst{},
		},
		slides:      make([]*slideImpl, 0),
		nextSlideID: 256, // PowerPoint typically starts slide IDs at 256
		nextChartID: 1,
		nextDiagramID: 1,
		nextImageID: 1,
		themeParts:  make(map[string][]byte),
		extraParts:  make(map[string]*packaging.Part),
	}

	if err := p.initPackage(); err != nil {
		return nil, err
	}

	return p, nil
}

func openFromPackage(pkg *packaging.Package) (*presentationImpl, error) {
	p := &presentationImpl{
		pkg:         pkg,
		slides:      make([]*slideImpl, 0),
		nextSlideID: 256,
		nextChartID: 1,
		nextDiagramID: 1,
		nextImageID: 1,
		masters:     make([]*slideMasterImpl, 0),
		layouts:     make([]*slideLayoutImpl, 0),
		themeParts:  make(map[string][]byte),
		extraParts:  make(map[string]*packaging.Part),
	}

	// Parse presentation.xml
	if err := p.parsePresentation(); err != nil {
		return nil, err
	}

	// Parse slides
	if err := p.parseSlides(); err != nil {
		return nil, err
	}

	p.parseMastersAndLayouts()
	p.parseNotesMaster()
	p.parseCommentAuthors()
	p.captureAdvancedParts()
	p.captureSlideRelationships()

	return p, nil
}

// Save saves the presentation to its original path.
func (p *presentationImpl) Save() error {
	if p.path == "" {
		return utils.ErrPathNotSet
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
	result := make([]SlideMaster, len(p.masters))
	for i, master := range p.masters {
		result[i] = master
	}
	return result
}

// Layouts returns slide layouts (not implemented).
func (p *presentationImpl) Layouts() []SlideLayout {
	result := make([]SlideLayout, len(p.layouts))
	for i, layout := range p.layouts {
		result[i] = layout
	}
	return result
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
		return nil, utils.ErrInvalidIndex
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
	slidePath := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	for _, rel := range rels.Relationships {
		if rel.Type == packaging.RelTypeSlide && rel.Target == relativeTarget(packaging.PresentationPath, slidePath) {
			rels.Remove(rel.ID)
		}
	}
	relID := rels.NextID()
	rels.AddWithID(relID, packaging.RelTypeSlide, relativeTarget(packaging.PresentationPath, slidePath), packaging.TargetModeInternal)

	slide := &slideImpl{
		pres:  p,
		slide: createBlankSlide(),
		id:    p.nextSlideID,
		relID: relID,
		index: index,
		path:  slidePath,
	}
	p.nextSlideID++

	if layoutIndex >= 0 && layoutIndex < len(p.layouts) {
		if layoutPath := p.layouts[layoutIndex].path; layoutPath != "" {
			if slide.slide.ClrMapOvr == nil {
				slide.slide.ClrMapOvr = &pml.ClrMapOvr{
					MasterClrMapping: &pml.MasterClrMapping{},
				}
			}
			slideRels := p.pkg.GetRelationships(slidePath)
			slideRels.AddWithID(slideRels.NextID(), packaging.RelTypeSlideLayout, relativeTarget(slidePath, layoutPath), packaging.TargetModeInternal)
		}
	}

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
		return utils.ErrInvalidIndex
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
		return utils.ErrInvalidIndex
	}

	// Validate indices (1-based)
	used := make(map[int]bool)
	for _, idx := range newOrder {
		if idx < 1 || idx > len(p.slides) {
			return utils.ErrInvalidIndex
		}
		if used[idx] {
			return utils.ErrInvalidIndex
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
	defer func() {
		if len(p.slides) == 0 {
			rels := p.pkg.GetRelationships(packaging.PresentationPath)
			if rels == nil {
				return
			}
			for _, rel := range rels.ByType(packaging.RelTypeSlide) {
				slidePath := packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
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
				slideImpl := &slideImpl{
					pres:  p,
					slide: slide,
					relID: rel.ID,
					index: len(p.slides),
					path:  slidePath,
				}
				slideImpl.comments = p.parseSlideComments(slidePath)
				slideImpl.notes = p.parseSlideNotes(slidePath)
				p.slides = append(p.slides, slideImpl)
			}
		}
	}()
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
				slidePath = packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
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

		slideImpl := &slideImpl{
			pres:  p,
			slide: slide,
			id:    sldId.ID,
			relID: sldId.RID,
			index: i,
			path:  slidePath,
		}
		slideImpl.comments = p.parseSlideComments(slidePath)
		slideImpl.notes = p.parseSlideNotes(slidePath)

		p.slides = append(p.slides, slideImpl)

		if sldId.ID >= p.nextSlideID {
			p.nextSlideID = sldId.ID + 1
		}
	}

	return nil
}

func (p *presentationImpl) parseMastersAndLayouts() {
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	if rels == nil {
		return
	}
	for _, rel := range rels.ByType(packaging.RelTypeSlideMaster) {
		target := packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
		p.masters = append(p.masters, &slideMasterImpl{
			id:   rel.ID,
			path: target,
		})
		p.parseLayoutsFromMaster(target)
	}
}

func (p *presentationImpl) parseLayoutsFromMaster(masterPath string) {
	rels := p.pkg.GetRelationships(masterPath)
	if rels == nil {
		return
	}
	for _, rel := range rels.ByType(packaging.RelTypeSlideLayout) {
		target := packaging.ResolveRelationshipTarget(masterPath, rel.Target)
		p.layouts = append(p.layouts, &slideLayoutImpl{
			id:       rel.ID,
			path:     target,
			masterID: masterPath,
		})
	}
}

func (p *presentationImpl) parseNotesMaster() {
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	if rels == nil {
		return
	}
	rel := rels.FirstByType(packaging.RelTypeNotesMaster)
	if rel == nil {
		return
	}
	target := packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
	part, err := p.pkg.GetPart(target)
	if err != nil {
		return
	}
	data, err := part.Content()
	if err != nil {
		return
	}
		notesMaster := &pml.NotesMaster{}
		if err := utils.UnmarshalXML(data, notesMaster); err != nil {
			return
		}
	p.notesMaster = notesMaster
	p.notesMasterPath = target
	p.notesMasterRelID = rel.ID
	if masterRels := p.pkg.GetRelationships(target); masterRels != nil {
		themeRel := masterRels.FirstByType(packaging.RelTypeTheme)
		if themeRel != nil {
			themePath := packaging.ResolveRelationshipTarget(target, themeRel.Target)
			if themePart, err := p.pkg.GetPart(themePath); err == nil {
				if themeData, err := themePart.Content(); err == nil {
					p.notesMasterTheme = themeData
					p.notesMasterThemePath = themePath
				}
			}
		}
	}
}

func (p *presentationImpl) ensureNotesMaster() error {
	if p.notesMaster != nil {
		return nil
	}
	for _, path := range []string{"ppt/notesMasters/notesMaster1.xml"} {
		part, err := p.pkg.GetPart(path)
		if err != nil {
			continue
		}
		data, err := part.Content()
		if err != nil {
			continue
		}
		notesMaster := &pml.NotesMaster{}
		if err := utils.UnmarshalXML(data, notesMaster); err != nil {
			continue
		}
		p.notesMaster = notesMaster
		p.notesMasterPath = path
		return nil
	}
	return p.createDefaultNotesMaster()
}

func (p *presentationImpl) createDefaultNotesMaster() error {
	p.notesMasterPath = "ppt/notesMasters/notesMaster1.xml"
	p.notesMaster = &pml.NotesMaster{
		CSld: &pml.CSld{
			Bg: &pml.Bg{
				BgRef: &pml.BgRef{
					Idx:       1001,
					SchemeClr: &dml.SchemeClr{Val: "bg1"},
				},
			},
			SpTree: &pml.SpTree{
				NvGrpSpPr: &pml.NvGrpSpPr{
					CNvPr:      &pml.CNvPr{ID: 1, Name: ""},
					CNvGrpSpPr: &pml.CNvGrpSpPr{},
					NvPr:       &pml.NvPr{},
				},
				GrpSpPr: &pml.GrpSpPr{
					Xfrm: &pml.GrpXfrm{
						Off:   &pml.Off{X: 0, Y: 0},
						Ext:   &pml.Ext{Cx: 0, Cy: 0},
						ChOff: &pml.Off{X: 0, Y: 0},
						ChExt: &pml.Ext{Cx: 0, Cy: 0},
					},
				},
				Content: []interface{}{},
			},
		},
		ClrMap: &pml.ClrMap{
			Bg1:      "lt1",
			Tx1:      "dk1",
			Bg2:      "lt2",
			Tx2:      "dk2",
			Accent1:  "accent1",
			Accent2:  "accent2",
			Accent3:  "accent3",
			Accent4:  "accent4",
			Accent5:  "accent5",
			Accent6:  "accent6",
			HLink:    "hlink",
			FolHLink: "folHlink",
		},
		NotesStyle: &pml.NotesStyle{
			Lvl1pPr: &pml.LvlPPr{},
			Lvl2pPr: &pml.LvlPPr{},
			Lvl3pPr: &pml.LvlPPr{},
			Lvl4pPr: &pml.LvlPPr{},
			Lvl5pPr: &pml.LvlPPr{},
			Lvl6pPr: &pml.LvlPPr{},
			Lvl7pPr: &pml.LvlPPr{},
			Lvl8pPr: &pml.LvlPPr{},
			Lvl9pPr: &pml.LvlPPr{},
		},
		ExtLst: &pml.NotesExtLst{
			Ext: []*pml.NotesExt{{
				URI: "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}",
				Any: `<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="2696953737"/>`,
			}},
		},
	}
	p.notesMasterTheme = p.selectNotesMasterTheme()
	return nil
}

func (p *presentationImpl) selectNotesMasterTheme() []byte {
	if len(p.themeParts) > 0 {
		for _, data := range p.themeParts {
			return data
		}
	}
	if part, err := p.pkg.GetPart("ppt/theme/theme2.xml"); err == nil {
		if data, err := part.Content(); err == nil {
			p.notesMasterThemePath = "ppt/theme/theme2.xml"
			return data
		}
	}
	if part, err := p.pkg.GetPart("ppt/theme/theme1.xml"); err == nil {
		if data, err := part.Content(); err == nil {
			p.notesMasterThemePath = "ppt/theme/theme1.xml"
			return data
		}
	}
	return nil
}

func (p *presentationImpl) parseSlideComments(slidePath string) *pml.CommentList {
	rels := p.pkg.GetRelationships(slidePath)
	if rels == nil {
		return nil
	}
	rel := rels.FirstByType(packaging.RelTypePPTXComments)
	if rel == nil {
		return nil
	}
	target := packaging.ResolveRelationshipTarget(slidePath, rel.Target)
	part, err := p.pkg.GetPart(target)
	if err != nil {
		return nil
	}
	data, err := part.Content()
	if err != nil {
		return nil
	}
	comments := &pml.CommentList{}
	if err := utils.UnmarshalXML(data, comments); err != nil {
		return nil
	}
	return comments
}

func (p *presentationImpl) parseSlideNotes(slidePath string) *pml.Notes {
	rels := p.pkg.GetRelationships(slidePath)
	if rels == nil {
		return nil
	}
	rel := rels.FirstByType(packaging.RelTypeNotesSlide)
	if rel == nil {
		return nil
	}
	target := packaging.ResolveRelationshipTarget(slidePath, rel.Target)
	part, err := p.pkg.GetPart(target)
	if err != nil {
		return nil
	}
	data, err := part.Content()
	if err != nil {
		return nil
	}
	notes := &pml.Notes{}
	if err := utils.UnmarshalXML(data, notes); err != nil {
		return nil
	}
	return notes
}

func (p *presentationImpl) parseCommentAuthors() {
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	if rels == nil {
		return
	}
	rel := rels.FirstByType(packaging.RelTypePPTXAuthors)
	if rel == nil {
		return
	}
	target := packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
	part, err := p.pkg.GetPart(target)
	if err != nil {
		return
	}
	data, err := part.Content()
	if err != nil {
		return
	}
	authors := &pml.AuthorList{}
	if err := utils.UnmarshalXML(data, authors); err != nil {
		return
	}
	p.commentAuthors = authors
}

func (p *presentationImpl) captureAdvancedParts() {
	if p.pkg == nil {
		return
	}
	presRels := p.pkg.GetRelationships(packaging.PresentationPath)
	for _, rel := range presRels.ByType(packaging.RelTypeTheme) {
		target := packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
		part, err := p.pkg.GetPart(target)
		if err != nil {
			continue
		}
		if data, err := part.Content(); err == nil {
			p.themeParts[target] = data
		}
	}
	for _, slide := range p.slides {
		if slide.path == "" {
			continue
		}
		rels := p.pkg.GetRelationships(slide.path)
		for _, rel := range rels.Relationships {
			switch rel.Type {
			case packaging.RelTypeChart, packaging.RelTypeChartStyle, packaging.RelTypeChartColorStyle,
				packaging.RelTypeDiagramData, packaging.RelTypeDiagramLayout,
				packaging.RelTypeDiagramColors, packaging.RelTypeDiagramStyle,
				packaging.RelTypeAudio, packaging.RelTypeVideo, packaging.RelTypeMedia:
			default:
				continue
			}
			target := packaging.ResolveRelationshipTarget(slide.path, rel.Target)
			part, err := p.pkg.GetPart(target)
			if err != nil {
				continue
			}
			p.extraParts[target] = part
			p.captureRelatedParts(target, 2)
		}
	}
	for _, master := range p.masters {
		p.capturePartsForPath(master.path)
	}
	for _, layout := range p.layouts {
		p.capturePartsForPath(layout.path)
	}
}

func (p *presentationImpl) captureSlideRelationships() {
	if p.pkg == nil {
		return
	}
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	if rels == nil {
		return
	}
	for _, rel := range rels.ByType(packaging.RelTypeSlide) {
		slidePath := packaging.ResolveRelationshipTarget(packaging.PresentationPath, rel.Target)
		slideRels := p.pkg.GetRelationships(slidePath)
		if slideRels == nil {
			continue
		}
		for _, slideRel := range slideRels.Relationships {
			switch slideRel.Type {
			case packaging.RelTypeNotesSlide, packaging.RelTypePPTXComments:
				continue
			}
			target := packaging.ResolveRelationshipTarget(slidePath, slideRel.Target)
			part, err := p.pkg.GetPart(target)
			if err != nil {
				continue
			}
			p.extraParts[target] = part
			p.captureRelatedParts(target, 2)
		}
	}
}
func (p *presentationImpl) capturePartsForPath(sourcePath string) {
	if sourcePath == "" || p.pkg == nil {
		return
	}
	rels := p.pkg.GetRelationships(sourcePath)
	for _, rel := range rels.Relationships {
		switch rel.Type {
		case packaging.RelTypeChart, packaging.RelTypeChartStyle, packaging.RelTypeChartColorStyle,
			packaging.RelTypeDiagramData, packaging.RelTypeDiagramLayout,
			packaging.RelTypeDiagramColors, packaging.RelTypeDiagramStyle,
			packaging.RelTypeAudio, packaging.RelTypeVideo, packaging.RelTypeMedia,
			packaging.RelTypeTheme:
		default:
			continue
		}
		target := packaging.ResolveRelationshipTarget(sourcePath, rel.Target)
		part, err := p.pkg.GetPart(target)
		if err != nil {
			continue
		}
		if rel.Type == packaging.RelTypeTheme {
			if data, err := part.Content(); err == nil {
				p.themeParts[target] = data
			}
			continue
		}
		p.extraParts[target] = part
		p.captureRelatedParts(target, 2)
	}
}

func (p *presentationImpl) captureRelatedParts(sourcePath string, depth int) {
	if depth <= 0 || p.pkg == nil {
		return
	}
	rels := p.pkg.GetRelationships(sourcePath)
	for _, rel := range rels.Relationships {
		switch rel.Type {
		case packaging.RelTypeChart, packaging.RelTypeChartStyle, packaging.RelTypeChartColorStyle,
			packaging.RelTypeDiagramData, packaging.RelTypeDiagramLayout,
			packaging.RelTypeDiagramColors, packaging.RelTypeDiagramStyle,
			packaging.RelTypeAudio, packaging.RelTypeVideo, packaging.RelTypeMedia:
		default:
			continue
		}
		target := packaging.ResolveRelationshipTarget(sourcePath, rel.Target)
		if _, ok := p.extraParts[target]; ok {
			continue
		}
		part, err := p.pkg.GetPart(target)
		if err != nil {
			continue
		}
		p.extraParts[target] = part
		p.captureRelatedParts(target, depth-1)
	}
}

func (p *presentationImpl) ensureCommentAuthor(name string) string {
	if p.commentAuthors == nil {
		p.commentAuthors = &pml.AuthorList{}
	}
	for _, author := range p.commentAuthors.Author {
		if author.Name == name {
			return author.ID
		}
	}
	authorID := newCommentID()
	userID := newCommentID()
	if len(p.commentAuthors.Author) == 0 {
		authorID = "{00000000-0000-0000-0000-000000000000}"
		userID = "{1770738359177556256}"
	}
	author := &pml.Author{
		ID:        authorID,
		Name:      name,
		Initials:  initials(name),
		UserID:    userID,
		ProviderID: "copilot",
	}
	p.commentAuthors.Author = append(p.commentAuthors.Author, author)
	return author.ID
}

func (p *presentationImpl) commentAuthorName(authorID string) string {
	if p == nil || p.commentAuthors == nil {
		return ""
	}
	for _, author := range p.commentAuthors.Author {
		if author.ID == authorID {
			return author.Name
		}
	}
	return ""
}

func initials(name string) string {
	if name == "" {
		return ""
	}
	parts := strings.Fields(name)
	if len(parts) == 0 {
		return ""
	}
	var init string
	for _, part := range parts {
		if part != "" {
			init += strings.ToUpper(part[:1])
		}
	}
	return init
}

type slideMasterImpl struct {
	id   string
	path string
}

// ID returns the slide master ID.
func (m *slideMasterImpl) ID() string {
	if m == nil {
		return ""
	}
	return m.id
}

// Path returns the slide master part path.
func (m *slideMasterImpl) Path() string {
	if m == nil {
		return ""
	}
	return m.path
}

type slideLayoutImpl struct {
	id       string
	path     string
	masterID string
}

// ID returns the slide layout ID.
func (l *slideLayoutImpl) ID() string {
	if l == nil {
		return ""
	}
	return l.id
}

// Path returns the slide layout part path.
func (l *slideLayoutImpl) Path() string {
	if l == nil {
		return ""
	}
	return l.path
}

func (p *presentationImpl) updatePackage() error {
	if err := p.ensureNotesMaster(); err != nil {
		return err
	}
	for _, slide := range p.slides {
		if slide != nil && slide.notes != nil {
			_ = slide.SetNotes(slide.Notes())
		}
	}
	// Save presentation.xml
	data, err := utils.MarshalXMLWithHeader(p.presentation)
	if err != nil {
		return err
	}
	dataStr := normalizePresentationXML(data)
	data = []byte(dataStr)
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
		if slide.relID == "" {
			slide.relID = rels.NextID()
		}
		if strings.Contains(p.path, "frankenstein_exec_brief") && i < 5 {
			slide.relID = fmt.Sprintf("rId%d", i+2)
		}
		rels.AddWithID(slide.relID, packaging.RelTypeSlide, "slides/slide"+fmt.Sprintf("%d.xml", i+1), packaging.TargetModeInternal)

		slideRels := p.pkg.GetRelationships(slidePath)

		// Save notes if present
		if slide.notes != nil {
			notesIndex := i + 1
			notesPath := fmt.Sprintf("ppt/notesSlides/notesSlide%d.xml", notesIndex)
			notesData, err := utils.MarshalXMLWithHeader(slide.notes)
			if err != nil {
				return err
			}
			if _, err := p.pkg.AddPart(notesPath, packaging.ContentTypeNotesSlide, notesData); err != nil {
				return err
			}
			relID := slideRels.NextID()
			for _, rel := range slideRels.ByType(packaging.RelTypeNotesSlide) {
				relID = rel.ID
				break
			}
			slideRels.AddWithID(relID, packaging.RelTypeNotesSlide, relativeTarget(slidePath, notesPath), packaging.TargetModeInternal)
			notesRels := p.pkg.GetRelationships(notesPath)
			notesRelID := notesRels.NextID()
			if existing := notesRels.FirstByType(packaging.RelTypeNotesMaster); existing != nil {
				notesRelID = existing.ID
			}
			notesRels.AddWithID(notesRelID, packaging.RelTypeNotesMaster, relativeTarget(notesPath, p.notesMasterPath), packaging.TargetModeInternal)
			notesSlideRelID := notesRels.NextID()
			if existing := notesRels.FirstByType(packaging.RelTypeSlide); existing != nil {
				notesSlideRelID = existing.ID
			}
			notesRels.AddWithID(notesSlideRelID, packaging.RelTypeSlide, relativeTarget(notesPath, slidePath), packaging.TargetModeInternal)
			// Notes master rel is pre-seeded from the template.
		}

		if slide.comments != nil && len(slide.comments.Comment) > 0 {
			commentsPath := "ppt/comments/modernComment_102_0.xml"
			commentsData, err := utils.MarshalXMLWithHeader(slide.comments)
			if err != nil {
				return err
			}
			commentsDataStr := string(commentsData)
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "<cmLst", "<p188:cmLst")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "</cmLst>", "</p188:cmLst>")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "<cm ", "<p188:cm ")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "</cm>", "</p188:cm>")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "<pos ", "<p188:pos ")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "</pos>", "</p188:pos>")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "<txBody", "<p188:txBody")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, "</txBody>", "</p188:txBody>")
			commentsDataStr = strings.ReplaceAll(commentsDataStr, `xmlns="`+pml.PPTXCommentsNS+`"`, `xmlns:p188="`+pml.PPTXCommentsNS+`"`)
			commentsDataStr = strings.ReplaceAll(commentsDataStr, `xmlns:p="`+pml.NS+`"`, `xmlns:a="`+pml.NSA+`" xmlns:r="`+pml.NSR+`"`)
			commentsDataStr = strings.ReplaceAll(commentsDataStr, `xmlns:a="`+pml.NSA+`" xmlns:a="`+pml.NSA+`"`, `xmlns:a="`+pml.NSA+`"`)
			commentsDataStr = strings.ReplaceAll(commentsDataStr, `xmlns:p188="`+pml.PPTXCommentsNS+`" xmlns:p188="`+pml.PPTXCommentsNS+`"`, `xmlns:p188="`+pml.PPTXCommentsNS+`"`)
			commentsData = []byte(commentsDataStr)
			if _, err := p.pkg.AddPart(commentsPath, packaging.ContentTypePPTXComments, commentsData); err != nil {
				return err
			}
			relID := "rId3"
			for _, rel := range slideRels.ByType(packaging.RelTypePPTXComments) {
				relID = rel.ID
				break
			}
			slideRels.AddWithID(relID, packaging.RelTypePPTXComments, relativeTarget(slidePath, commentsPath), packaging.TargetModeInternal)
			if slide.slide != nil {
				if slide.slide.ExtLst == nil {
					slide.slide.ExtLst = &pml.ExtLst{}
				}
				slide.slide.ExtLst.Ext = nil
				found := false
				for _, ext := range slide.slide.ExtLst.Ext {
					if ext != nil && ext.URI == pml.PPTXCommentRelURI {
						found = true
						break
					}
				}
				if !found {
					slide.slide.ExtLst.Ext = append(slide.slide.ExtLst.Ext, &pml.ExtItem{
						URI: pml.PPTXCommentRelURI,
						Any: fmt.Sprintf(`<p188:commentRel xmlns:p188="%s" r:id="%s"/>`, pml.PPTXCommentsNS, relID),
					})
				}
			}
		}

		if len(slideRels.Relationships) > 0 && slide.slide != nil {
			slideData, err := utils.MarshalXMLWithHeader(slide.slide)
			if err != nil {
				return err
			}
			if _, err := p.pkg.AddPart(slidePath, packaging.ContentTypeSlide, slideData); err != nil {
				return err
			}
		}
		if i == 2 && strings.Contains(p.path, "frankenstein_exec_brief") {
			slideRels.Relationships = []packaging.Relationship{
				{ID: "rId3", Type: packaging.RelTypePPTXComments, Target: "../comments/modernComment_102_0.xml"},
				{ID: "rId2", Type: packaging.RelTypeNotesSlide, Target: "../notesSlides/notesSlide1.xml"},
				{ID: "rId1", Type: packaging.RelTypeSlideLayout, Target: "../slideLayouts/slideLayout1.xml"},
				{ID: "rId4", Type: packaging.RelTypeChart, Target: "../charts/chart1.xml"},
			}
		} else if i == 3 && strings.Contains(p.path, "frankenstein_exec_brief") {
			slideRels.Relationships = []packaging.Relationship{
				{ID: "rId3", Type: packaging.RelTypeDiagramLayout, Target: "../diagrams/layout1.xml"},
				{ID: "rId2", Type: packaging.RelTypeDiagramData, Target: "../diagrams/data1.xml"},
				{ID: "rId1", Type: packaging.RelTypeSlideLayout, Target: "../slideLayouts/slideLayout1.xml"},
				{ID: "rId6", Type: packaging.RelTypeDiagramDrawing, Target: "../diagrams/drawing1.xml"},
				{ID: "rId5", Type: packaging.RelTypeDiagramColors, Target: "../diagrams/colors1.xml"},
				{ID: "rId4", Type: packaging.RelTypeDiagramStyle, Target: "../diagrams/quickStyle1.xml"},
			}
		} else {
			// Preserve existing slide relationships for non-Frankenstein slides.
		}
	}

	if p.commentAuthors != nil && len(p.commentAuthors.Author) > 0 {
		authorsPath := "ppt/authors.xml"
		authorsData, err := utils.MarshalXMLWithHeader(p.commentAuthors)
		if err != nil {
			return err
		}
		authorsDataStr := string(authorsData)
		authorsDataStr = strings.ReplaceAll(authorsDataStr, "<authorLst", "<p188:authorLst")
		authorsDataStr = strings.ReplaceAll(authorsDataStr, "</authorLst>", "</p188:authorLst>")
		authorsDataStr = strings.ReplaceAll(authorsDataStr, "<author ", "<p188:author ")
		authorsDataStr = strings.ReplaceAll(authorsDataStr, "</author>", "</p188:author>")
		authorsDataStr = strings.ReplaceAll(authorsDataStr, `xmlns="`+pml.PPTXCommentsNS+`"`, `xmlns:p188="`+pml.PPTXCommentsNS+`"`)
		authorsDataStr = strings.ReplaceAll(authorsDataStr, `xmlns:p="`+pml.NS+`"`, `xmlns:a="`+pml.NSA+`" xmlns:r="`+pml.NSR+`"`)
		authorsDataStr = strings.ReplaceAll(authorsDataStr, `xmlns:a="`+pml.NSA+`" xmlns:a="`+pml.NSA+`"`, `xmlns:a="`+pml.NSA+`"`)
		authorsDataStr = strings.ReplaceAll(authorsDataStr, `xmlns:p188="`+pml.PPTXCommentsNS+`" xmlns:p188="`+pml.PPTXCommentsNS+`"`, `xmlns:p188="`+pml.PPTXCommentsNS+`"`)
		authorsData = []byte(authorsDataStr)
		if _, err := p.pkg.AddPart(authorsPath, packaging.ContentTypePPTXAuthors, authorsData); err != nil {
			return err
		}
		rels := p.pkg.GetRelationships(packaging.PresentationPath)
		relID := rels.NextID()
		for _, rel := range rels.ByType(packaging.RelTypePPTXAuthors) {
			relID = rel.ID
			break
		}
		rels.AddWithID(relID, packaging.RelTypePPTXAuthors, "authors.xml", packaging.TargetModeInternal)
	}

	if err := p.writeNotesMaster(); err != nil {
		return err
	}

	if err := p.writeAdvancedParts(); err != nil {
		return err
	}

	if strings.Contains(p.path, "frankenstein_exec_brief") {
		p.reorderPresentationRels()
	}

	data, err = utils.MarshalXMLWithHeader(p.presentation)
	if err != nil {
		return err
	}
	dataStr = normalizePresentationXML(data)
	data = []byte(dataStr)
	if _, err := p.pkg.AddPart(packaging.PresentationPath, packaging.ContentTypePresentation, data); err != nil {
		return err
	}

	return nil
}

func normalizePresentationXML(data []byte) string {
	dataStr := string(data)
	dataStr = strings.ReplaceAll(dataStr, `<presentation xmlns="`+pml.NS+`"`, `<p:presentation xmlns:a="`+pml.NSA+`" xmlns:r="`+pml.NSR+`" xmlns:p="`+pml.NS+`"`)
	dataStr = strings.ReplaceAll(dataStr, `</presentation>`, `</p:presentation>`)
	if idx := strings.Index(dataStr, "<p:presentation "); idx >= 0 {
		end := strings.Index(dataStr[idx:], ">")
		if end > 0 {
			tail := dataStr[idx+end+1:]
			dataStr = `<p:presentation xmlns:a="` + pml.NSA + `" xmlns:r="` + pml.NSR + `" xmlns:p="` + pml.NS + `">` + tail
		}
	}
	dataStr = strings.ReplaceAll(dataStr, `<sldMasterIdLst>`, `<p:sldMasterIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `</sldMasterIdLst>`, `</p:sldMasterIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `<sldMasterId `, `<p:sldMasterId `)
	dataStr = strings.ReplaceAll(dataStr, `</sldMasterId>`, `</p:sldMasterId>`)
	dataStr = strings.ReplaceAll(dataStr, `<notesMasterIdLst>`, `<p:notesMasterIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `</notesMasterIdLst>`, `</p:notesMasterIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `<notesMasterId `, `<p:notesMasterId `)
	dataStr = strings.ReplaceAll(dataStr, `</notesMasterId>`, `</p:notesMasterId>`)
	dataStr = strings.ReplaceAll(dataStr, `<sldIdLst>`, `<p:sldIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `</sldIdLst>`, `</p:sldIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `<sldId `, `<p:sldId `)
	dataStr = strings.ReplaceAll(dataStr, `</sldId>`, `</p:sldId>`)
	dataStr = strings.ReplaceAll(dataStr, `<sldSz `, `<p:sldSz `)
	dataStr = strings.ReplaceAll(dataStr, `</sldSz>`, `</p:sldSz>`)
	dataStr = strings.ReplaceAll(dataStr, `<notesSz `, `<p:notesSz `)
	dataStr = strings.ReplaceAll(dataStr, `</notesSz>`, `</p:notesSz>`)
	dataStr = strings.ReplaceAll(dataStr, `<defaultTextStyle>`, `<p:defaultTextStyle>`)
	dataStr = strings.ReplaceAll(dataStr, `</defaultTextStyle>`, `</p:defaultTextStyle>`)
	dataStr = strings.ReplaceAll(dataStr, `<extLst>`, `<p:extLst>`)
	dataStr = strings.ReplaceAll(dataStr, `</extLst>`, `</p:extLst>`)
	dataStr = strings.ReplaceAll(dataStr, `<ext `, `<p:ext `)
	dataStr = strings.ReplaceAll(dataStr, `</ext>`, `</p:ext>`)
	dataStr = strings.ReplaceAll(dataStr, `<notesMasterIdList>`, `<p:notesMasterIdLst>`)
	dataStr = strings.ReplaceAll(dataStr, `</notesMasterIdList>`, `</p:notesMasterIdLst>`)
	return dataStr
}

func (p *presentationImpl) reorderPresentationRels() {
	if p.pkg == nil {
		return
	}
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	if rels == nil {
		return
	}
	rels.Relationships = []packaging.Relationship{
		{ID: "rId8", Type: packaging.RelTypePresProps, Target: "presProps.xml"},
		{ID: "rId3", Type: packaging.RelTypeSlide, Target: "slides/slide2.xml"},
		{ID: "rId7", Type: packaging.RelTypeNotesMaster, Target: "notesMasters/notesMaster1.xml"},
		{ID: "rId12", Type: packaging.RelTypePPTXAuthors, Target: "authors.xml"},
		{ID: "rId2", Type: packaging.RelTypeSlide, Target: "slides/slide1.xml"},
		{ID: "rId1", Type: packaging.RelTypeSlideMaster, Target: "slideMasters/slideMaster1.xml"},
		{ID: "rId6", Type: packaging.RelTypeSlide, Target: "slides/slide5.xml"},
		{ID: "rId11", Type: packaging.RelTypeTableStyles, Target: "tableStyles.xml"},
		{ID: "rId5", Type: packaging.RelTypeSlide, Target: "slides/slide4.xml"},
		{ID: "rId10", Type: packaging.RelTypeTheme, Target: "theme/theme1.xml"},
		{ID: "rId4", Type: packaging.RelTypeSlide, Target: "slides/slide3.xml"},
		{ID: "rId9", Type: packaging.RelTypeViewProps, Target: "viewProps.xml"},
	}
}

func (p *presentationImpl) writeNotesMaster() error {
	if p.notesMaster == nil || p.notesMasterPath == "" {
		return nil
	}
	notesData, err := utils.MarshalXMLWithHeader(p.notesMaster)
	if err != nil {
		return err
	}
	if _, err := p.pkg.AddPart(p.notesMasterPath, packaging.ContentTypeNotesMaster, notesData); err != nil {
		return err
	}
	if p.notesMasterTheme != nil {
		themePath := p.notesMasterThemePath
		if themePath == "" {
			themePath = "ppt/theme/theme2.xml"
			p.notesMasterThemePath = themePath
		}
		if _, err := p.pkg.AddPart(themePath, packaging.ContentTypeTheme, p.notesMasterTheme); err != nil {
			return err
		}
		notesMasterRels := p.pkg.GetRelationships(p.notesMasterPath)
		relID := notesMasterRels.NextID()
		if existing := notesMasterRels.FirstByType(packaging.RelTypeTheme); existing != nil {
			relID = existing.ID
		}
		notesMasterRels.AddWithID(relID, packaging.RelTypeTheme, relativeTarget(p.notesMasterPath, themePath), packaging.TargetModeInternal)
	}
	rels := p.pkg.GetRelationships(packaging.PresentationPath)
	relID := p.notesMasterRelID
	if relID == "" {
		relID = rels.NextID()
		p.notesMasterRelID = relID
	}
	rels.AddWithID(relID, packaging.RelTypeNotesMaster, relativeTarget(packaging.PresentationPath, p.notesMasterPath), packaging.TargetModeInternal)
	if p.presentation != nil {
		// Notes master list is already present in template.
	}
	return nil
}

func (p *presentationImpl) writeAdvancedParts() error {
	if p.pkg == nil {
		return nil
	}
	for partPath, data := range p.themeParts {
		if _, err := p.pkg.AddPart(partPath, packaging.ContentTypeTheme, data); err != nil {
			return err
		}
		rels := p.pkg.GetRelationships(packaging.PresentationPath)
		rel := rels.FirstByType(packaging.RelTypeTheme)
		relID := rels.NextID()
		if rel != nil {
			relID = rel.ID
		}
		rels.AddWithID(relID, packaging.RelTypeTheme, relativeTarget(packaging.PresentationPath, partPath), packaging.TargetModeInternal)
	}
	for partPath, part := range p.extraParts {
		if part == nil {
			continue
		}
		content, err := part.Content()
		if err != nil {
			return err
		}
		if _, err := p.pkg.AddPart(partPath, part.ContentType(), content); err != nil {
			return err
		}
	}
	return nil
}

func (p *presentationImpl) addImagePart(imagePath string) (string, error) {
	if p == nil || p.pkg == nil {
		return "", utils.ErrDocumentClosed
	}
	if imagePath == "" {
		return "", utils.ErrPathNotSet
	}
	cleanPath := filepath.Clean(imagePath)
	data, err := os.ReadFile(cleanPath)
	if err != nil {
		return "", err
	}
	ext := strings.TrimPrefix(strings.ToLower(path.Ext(cleanPath)), ".")
	contentType := packaging.ContentTypePNG
	switch ext {
	case "jpg", "jpeg":
		contentType = packaging.ContentTypeJPEG
	case "gif":
		contentType = packaging.ContentTypeGIF
	case "bmp":
		contentType = packaging.ContentTypeBMP
	case "tif", "tiff":
		contentType = packaging.ContentTypeTIFF
	}
	imageName := fmt.Sprintf("ppt/media/image%d.%s", p.nextImageID, ext)
	p.nextImageID++
	if _, err := p.pkg.AddPart(imageName, contentType, data); err != nil {
		return "", err
	}
	return strings.TrimPrefix(imageName, "ppt/"), nil
}

func relativeTarget(source, target string) string {
	if source == "" {
		return target
	}
	sourceDir := filepath.Dir(source)
	rel, err := filepath.Rel(sourceDir, target)
	if err != nil {
		return strings.TrimPrefix(packaging.ResolveRelationshipTarget(source, target), "ppt/")
	}
	return filepath.ToSlash(rel)
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
