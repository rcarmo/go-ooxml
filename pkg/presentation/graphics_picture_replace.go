package presentation

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type PictureReplacementOptions struct {
	ContentType string `json:"contentType"`
}
type PictureReplacementReceipt struct {
	PictureReceipt
	PreviousRelationshipID string `json:"previousRelationshipId"`
	PreviousMediaPart      string `json:"previousMediaPart"`
}

// ReplacePicture isolates one embedded-only dependency and retains prior media.
func (s *EditSession) ReplacePicture(part string, id uint32, payload []byte, options PictureReplacementOptions) (PictureReplacementReceipt, error) {
	var none PictureReplacementReceipt
	if id == 0 || id > math.MaxInt32 || len(payload) == 0 || len(payload) > 64*1024*1024 {
		return none, graphicsUnsupported("replacement input bounds")
	}
	png := len(payload) >= 8 && bytes.Equal(payload[:8], []byte{137, 80, 78, 71, 13, 10, 26, 10})
	jpeg := len(payload) >= 4 && payload[0] == 255 && payload[1] == 216 && payload[len(payload)-2] == 255 && payload[len(payload)-1] == 217
	if !(options.ContentType == "image/png" && png || options.ContentType == "image/jpeg" && jpeg) {
		return none, graphicsUnsupported("replacement MIME/signature")
	}
	if !s.slides[part] {
		return none, graphicsUnsupported("slide not enrolled")
	}
	if err := s.manipulationProtection(); err != nil {
		return none, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	pictures, err := s.InspectPictures(part)
	if err != nil {
		return none, err
	}
	var selected *PictureInfo
	for i := range pictures {
		if pictures[i].ShapeID == id {
			selected = &pictures[i]
		}
	}
	if selected == nil {
		return none, editRefusal("PPTX_PICTURE_NOT_FOUND", "exact picture identity absent")
	}
	if selected.Embedded == nil || selected.Linked != nil {
		return none, graphicsUnsupported("embedded-only picture required")
	}
	source, hash, err := s.pkg.Part(part)
	if err != nil {
		return none, err
	}
	doc, err := losslessxml.Parse(source)
	if err != nil {
		return none, err
	}
	var blip losslessxml.Element
	for _, node := range doc.Elements() {
		if node.Name() != name(packaging.NSPresentationML, "pic") {
			continue
		}
		got, _, _, e := graphicsIdentity(doc, node, "nvPicPr")
		if e != nil {
			return none, e
		}
		if got == id {
			fill, e := graphicsOne(doc, node, packaging.NSPresentationML, "blipFill", false)
			if e != nil {
				return none, e
			}
			blip, e = graphicsOne(doc, fill, packaging.NSDrawingML, "blip", false)
			if e != nil {
				return none, e
			}
			break
		}
	}
	for _, node := range doc.Elements() {
		if node == blip || !manipulationWithin(node, blip) {
			continue
		}
		for _, a := range node.Attributes() {
			if a.Name.Space == packaging.NSDocumentRelationships && (a.Name.Local == "embed" || a.Name.Local == "link") {
				return none, graphicsUnsupported("paired image extension")
			}
		}
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return none, err
	}
	used := map[string]bool{}
	rids := map[string]bool{}
	for _, p := range graph.Parts {
		used[strings.ToLower(p.Name)] = true
	}
	for _, e := range graph.Edges {
		if e.Source == part {
			rids[e.ID] = true
		}
	}
	ext := "png"
	if options.ContentType == "image/jpeg" {
		ext = "jpeg"
	}
	media := ""
	for i := 1; i <= len(used)+1; i++ {
		candidate := fmt.Sprintf("ppt/media/image%d.%s", i, ext)
		if !used[strings.ToLower(candidate)] {
			media = candidate
			break
		}
	}
	if media == "" {
		return none, graphicsUnsupported("no replacement media identity")
	}
	rid := ""
	for i := 1; i <= len(rids)+1; i++ {
		candidate := "rId" + strconv.Itoa(i)
		if !rids[candidate] {
			rid = candidate
			break
		}
	}
	next, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: blip, Name: xml.Name{Space: packaging.NSDocumentRelationships, Local: "embed"}, Value: rid}})
	if err != nil {
		return none, graphicsUnsupported("embed value lexical update")
	}
	plan, err := s.pkg.PlanGraphMutation(packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: media, ContentType: options.ContentType, Data: payload}}, Relationships: []packaging.RelationshipAddition{{Source: part, ID: rid, Type: packaging.RelTypeImage, TargetPart: media}}, Replacements: []packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: next}}})
	if err != nil {
		return none, err
	}
	if err = s.pkg.ApplyGraphPlan(plan); err != nil {
		return none, err
	}
	s.generation++
	return PictureReplacementReceipt{PictureReceipt{id, part, media, rid}, selected.Embedded.RelationshipID, selected.Embedded.PartName}, nil
}
