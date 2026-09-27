package spreadsheet

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const xdrNS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

// ImageTarget is a session-bound loaded picture selection. It is not a legacy drawing handle.
type ImageTarget struct {
	session                      *EditSession
	generation                   uint64
	sheet                        string
	shapeID                      uint32
	drawing, relationship, media string
	consumed                     bool
}

func imageAttr(e losslessxml.Element, name xml.Name) string {
	for _, a := range e.Attributes() {
		if a.Name == name {
			return a.Value
		}
	}
	return ""
}
func imageChildren(doc *losslessxml.Document, parent losslessxml.Element, name xml.Name) []losslessxml.Element {
	var out []losslessxml.Element
	for _, e := range doc.Elements() {
		if p, ok := e.Parent(); ok && p == parent && e.Name() == name {
			out = append(out, e)
		}
	}
	return out
}
func imageProtection(doc *losslessxml.Document) error {
	for _, e := range doc.Elements() {
		n := e.Name()
		if n.Space != packaging.NSSpreadsheetML {
			return editRefusal("unsupported_structure", "worksheet/workbook extension prevents image ownership proof")
		}
		switch n.Local {
		case "workbookProtection", "sheetProtection":
			text, leaf := e.Text()
			if len(e.Attributes()) > 0 || !leaf || strings.TrimSpace(text) != "" {
				return editRefusal("protected_operation", "image replacement on protected content")
			}
		case "extLst", "AlternateContent":
			return editRefusal("unsupported_structure", "extended image ownership not supported")
		}
	}
	return nil
}

// FindImage selects one ordinary directly anchored picture by sheet and cNvPr ID.
// The drawing must have one inbound edge and its image relationship exactly one
// XML consumer. Distinct edges may share media; the original media is retained.
func (s *EditSession) FindImage(sheet string, shapeID uint32) (*ImageTarget, error) {
	if shapeID == 0 {
		return nil, editRefusal("missing_target", "nonzero picture ID required")
	}
	sheetPart, ok := s.sheets[sheet]
	if !ok {
		return nil, editRefusal("missing_target", "worksheet absent")
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return nil, err
	}
	parts := map[string]packaging.GraphPart{}
	for _, p := range graph.Parts {
		parts[p.Name] = p
	}
	if parts[sheetPart].Inbound != 1 {
		return nil, editRefusal("ambiguous_target", "worksheet part has shared package ownership")
	}
	main, _, err := s.pkg.Part(s.main)
	if err != nil {
		return nil, err
	}
	mainDoc, err := losslessxml.Parse(main)
	if err != nil {
		return nil, err
	}
	if err = imageProtection(mainDoc); err != nil {
		return nil, err
	}
	sheetData, _, err := s.pkg.Part(sheetPart)
	if err != nil {
		return nil, err
	}
	sheetDoc, err := losslessxml.Parse(sheetData)
	if err != nil {
		return nil, err
	}
	if err = imageProtection(sheetDoc); err != nil {
		return nil, err
	}
	root := sheetDoc.Elements()[0]
	if root.Name() != expanded("worksheet") {
		return nil, editRefusal("unsupported_structure", "worksheet root required")
	}
	drawings := imageChildren(sheetDoc, root, expanded("drawing"))
	if len(drawings) != 1 {
		return nil, editRefusal("ambiguous_target", "exactly one direct drawing required")
	}
	rid := imageAttr(drawings[0], xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"})
	if rid == "" {
		return nil, editRefusal("missing_target", "drawing relationship ID absent")
	}
	refs := 0
	for _, e := range sheetDoc.Elements() {
		for _, a := range e.Attributes() {
			if a.Name.Space == packaging.NSDocumentRelationships && a.Value == rid {
				refs++
			}
		}
	}
	if refs != 1 {
		return nil, editRefusal("ambiguous_target", "drawing edge has multiple XML consumers")
	}
	drawing := ""
	for _, e := range graph.Edges {
		if e.Source == sheetPart && e.ID == rid {
			if e.Type != packaging.RelTypeDrawing || e.External || strings.Contains(e.Target, "#") {
				return nil, editRefusal("unsupported_structure", "drawing relationship type or mode unsupported")
			}
			drawing = e.ResolvedPart
		}
	}
	if drawing == "" || parts[drawing].ContentType != packaging.ContentTypeDrawing || parts[drawing].Inbound != 1 {
		return nil, editRefusal("ambiguous_target", "drawing absent, shared or wrong type")
	}
	data, _, err := s.pkg.Part(drawing)
	if err != nil {
		return nil, err
	}
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	elements := doc.Elements()
	if elements[0].Name() != (xml.Name{Space: xdrNS, Local: "wsDr"}) {
		return nil, editRefusal("unsupported_structure", "drawing root required")
	}
	seen := map[uint64]bool{}
	var identity losslessxml.Element
	found := false
	for _, e := range elements {
		n := e.Name()
		if n.Space != xdrNS && n.Space != packaging.NSDrawingML {
			return nil, editRefusal("unsupported_structure", "unknown drawing namespace")
		}
		if n.Local == "extLst" || n.Local == "picLocks" || n.Local == "grpSp" || n.Local == "graphicFrame" {
			return nil, editRefusal("unsupported_structure", "group/chart/extension/locked drawing unsupported")
		}
		for _, a := range e.Attributes() {
			if a.Name.Space != "" && a.Name.Space != packaging.NSDocumentRelationships && a.Name.Space != "http://www.w3.org/XML/1998/namespace" {
				return nil, editRefusal("unsupported_structure", "unknown drawing attribute namespace")
			}
		}
		if n == (xml.Name{Space: xdrNS, Local: "cNvPr"}) {
			id, err := strconv.ParseUint(attr(e, "id"), 10, 32)
			if err != nil || id == 0 || seen[id] {
				return nil, editRefusal("ambiguous_target", "invalid or duplicate shape ID")
			}
			seen[id] = true
			if id == uint64(shapeID) {
				identity = e
				found = true
			}
		}
	}
	if !found {
		return nil, editRefusal("missing_target", "picture ID absent")
	}
	nv, ok := identity.Parent()
	if !ok || nv.Name() != (xml.Name{Space: xdrNS, Local: "nvPicPr"}) {
		return nil, editRefusal("unsupported_structure", "selected shape is not a picture")
	}
	pic, ok := nv.Parent()
	if !ok || pic.Name() != (xml.Name{Space: xdrNS, Local: "pic"}) {
		return nil, editRefusal("unsupported_structure", "picture owner invalid")
	}
	anchor, ok := pic.Parent()
	if !ok || anchor.Name().Space != xdrNS {
		return nil, editRefusal("unsupported_structure", "picture anchor absent")
	}
	switch anchor.Name().Local {
	case "oneCellAnchor", "twoCellAnchor", "absoluteAnchor":
	default:
		return nil, editRefusal("unsupported_structure", "unsupported picture anchor")
	}
	owner, ok := anchor.Parent()
	if !ok || owner != elements[0] {
		return nil, editRefusal("unsupported_structure", "nested picture anchor")
	}
	fills := imageChildren(doc, pic, xml.Name{Space: xdrNS, Local: "blipFill"})
	if len(fills) != 1 {
		return nil, editRefusal("ambiguous_target", "picture fill ambiguous")
	}
	blips := imageChildren(doc, fills[0], xml.Name{Space: packaging.NSDrawingML, Local: "blip"})
	if len(blips) != 1 {
		return nil, editRefusal("ambiguous_target", "picture image ambiguous")
	}
	blip := blips[0]
	imageRID := imageAttr(blip, xml.Name{Space: packaging.NSDocumentRelationships, Local: "embed"})
	if imageRID == "" || imageAttr(blip, xml.Name{Space: packaging.NSDocumentRelationships, Local: "link"}) != "" {
		return nil, editRefusal("unsupported_structure", "linked picture unsupported")
	}
	uses := 0
	for _, e := range elements {
		for _, a := range e.Attributes() {
			if a.Name.Space == packaging.NSDocumentRelationships && a.Value == imageRID {
				uses++
			}
		}
	}
	if uses != 1 {
		return nil, editRefusal("ambiguous_target", "image relationship is shared by XML consumers")
	}
	media := ""
	for _, e := range graph.Edges {
		if e.Source == drawing && e.ID == imageRID {
			if e.Type != packaging.RelTypeImage || e.External || strings.Contains(e.Target, "#") {
				return nil, editRefusal("unsupported_structure", "picture relationship unsupported")
			}
			media = e.ResolvedPart
		}
	}
	if media == "" || (parts[media].ContentType != packaging.ContentTypePNG && parts[media].ContentType != packaging.ContentTypeJPEG) {
		return nil, editRefusal("unsupported_structure", "loaded picture must be PNG or JPEG")
	}
	return &ImageTarget{session: s, generation: s.generation, sheet: sheet, shapeID: shapeID, drawing: drawing, relationship: imageRID, media: media}, nil
}

// ReplaceImage validates a bounded PNG/JPEG, allocates a fresh media member and
// retargets only the selected edge. No drawing geometry/crop/cache XML is changed.
// Identical image bytes are a no-op and leave the handle usable.
func (s *EditSession) ReplaceImage(target *ImageTarget, data []byte) error {
	if target == nil || target.session != s || target.consumed || target.generation != s.generation {
		return editRefusal("stale_target", "foreign, consumed or stale image target")
	}
	current, err := s.FindImage(target.sheet, target.shapeID)
	if err != nil {
		return err
	}
	if current.drawing != target.drawing || current.relationship != target.relationship || current.media != target.media {
		return editRefusal("stale_target", "picture ownership changed")
	}
	if len(data) == 0 || len(data) > 32<<20 {
		return editRefusal("unsupported_structure", "replacement image must be between 1 and 32 MiB bytes")
	}
	config, format, err := image.DecodeConfig(bytes.NewReader(data))
	if err != nil || format != "png" && format != "jpeg" || config.Width <= 0 || config.Height <= 0 || int64(config.Width)*int64(config.Height) > 16_000_000 {
		return editRefusal("unsupported_structure", "invalid, unsupported or excessive replacement image")
	}
	_, decoded, err := image.Decode(bytes.NewReader(data))
	if err != nil || decoded != format {
		return editRefusal("unsupported_structure", "replacement image decode failed")
	}
	original, _, err := s.pkg.Part(target.media)
	if err != nil {
		return err
	}
	if bytes.Equal(data, original) {
		return nil
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	taken := map[string]bool{}
	for _, p := range graph.Parts {
		taken[strings.ToLower(p.Name)] = true
	}
	extension, mime := "png", packaging.ContentTypePNG
	if format == "jpeg" {
		extension, mime = "jpg", packaging.ContentTypeJPEG
	}
	name := ""
	for i := 1; i <= len(taken)+1; i++ {
		candidate := fmt.Sprintf("xl/media/image%d.%s", i, extension)
		if !taken[strings.ToLower(candidate)] {
			name = candidate
			break
		}
	}
	if name == "" {
		return editRefusal("ambiguous_target", "media name allocation failed")
	}
	plan, err := s.pkg.PlanGraphMutation(packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: name, ContentType: mime, Data: data}}, Retargets: []packaging.RelationshipRetarget{{Source: target.drawing, ID: target.relationship, TargetPart: name}}})
	if err != nil {
		return err
	}
	if err = s.pkg.ApplyGraphPlan(plan); err != nil {
		return err
	}
	s.generation++
	target.consumed = true
	return nil
}
