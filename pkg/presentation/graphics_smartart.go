package presentation

import (
	"math"
	"sort"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const graphicsDiagramNS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
const graphicsPersistDiagramNS = "http://schemas.microsoft.com/office/drawing/2008/diagram"
const graphicsDiagramDrawingRel = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
const graphicsRelationNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

type SmartArtPart struct {
	PartName    string `json:"partName"`
	ContentType string `json:"contentType"`
	ByteLength  int    `json:"byteLength"`
}
type SmartArtRoot struct {
	SmartArtPart
	Role           string `json:"role"`
	RelationshipID string `json:"relationshipId"`
	Target         string `json:"target"`
}
type SmartArtEdge struct {
	Owner          string  `json:"owner"`
	RelationshipID string  `json:"relationshipId"`
	Type           string  `json:"type"`
	Target         string  `json:"target"`
	External       bool    `json:"external"`
	PartName       *string `json:"partName"`
}
type SmartArtLimits struct {
	DataEditing         bool   `json:"dataEditing"`
	LayoutEvaluation    bool   `json:"layoutEvaluation"`
	StyleEvaluation     bool   `json:"styleEvaluation"`
	DrawingRegeneration bool   `json:"drawingRegeneration"`
	Reason              string `json:"reason"`
}
type SmartArtInfo struct {
	ShapeID      uint32         `json:"shapeId"`
	Name         string         `json:"name"`
	SlidePart    string         `json:"slidePart"`
	Roots        []SmartArtRoot `json:"roots"`
	Parts        []SmartArtPart `json:"parts"`
	Edges        []SmartArtEdge `json:"edges"`
	DrawingParts []string       `json:"drawingParts"`
	Limits       SmartArtLimits `json:"limits"`
}

func graphicsSmartArtUnsupported(detail string) error {
	return editRefusal("PPTX_SMARTART_UNSUPPORTED", detail)
}
func graphicsSmartOne(doc *losslessxml.Document, parent losslessxml.Element, ns, local string) (losslessxml.Element, error) {
	n, e := graphicsOne(doc, parent, ns, local, false)
	if e != nil {
		return n, graphicsSmartArtUnsupported("unique " + local + " required")
	}
	return n, nil
}
func (s *EditSession) InspectSmartArt(part string) ([]SmartArtInfo, error) {
	if !s.slides[part] {
		return nil, graphicsSmartArtUnsupported("slide not enrolled")
	}
	b, _, e := s.pkg.Part(part)
	if e != nil {
		return nil, e
	}
	doc, e := losslessxml.Parse(b)
	if e != nil {
		return nil, e
	}
	es := doc.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, graphicsSmartArtUnsupported("slide root")
	}
	common, e := graphicsSmartOne(doc, es[0], packaging.NSPresentationML, "cSld")
	if e != nil {
		return nil, e
	}
	tree, e := graphicsSmartOne(doc, common, packaging.NSPresentationML, "spTree")
	if e != nil {
		return nil, e
	}
	ids := map[int64]bool{}
	for _, n := range es {
		if n.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		raw, _ := graphicsAttr(n, "", "id")
		v, er := graphicsNumber(raw, 1, math.MaxInt32)
		if er != nil || strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") || ids[v] {
			return nil, graphicsSmartArtUnsupported("duplicate/invalid slide identity")
		}
		ids[v] = true
	}
	graph, e := s.pkg.Graph()
	if e != nil {
		return nil, e
	}
	types := map[string]string{}
	byOwner := map[string][]packaging.Edge{}
	for _, p := range graph.Parts {
		types[p.Name] = p.ContentType
	}
	for _, edge := range graph.Edges {
		byOwner[edge.Source] = append(byOwner[edge.Source], edge)
	}
	asset := func(p string) (SmartArtPart, error) {
		b, _, er := s.pkg.Part(p)
		if er != nil || types[p] == "" {
			return SmartArtPart{}, graphicsSmartArtUnsupported("missing typed dependency")
		}
		return SmartArtPart{p, types[p], len(b)}, nil
	}
	root := func(p, mime, local, ns string) error {
		if types[p] != mime {
			return graphicsSmartArtUnsupported("dependency MIME")
		}
		raw, _, er := s.pkg.Part(p)
		if er != nil {
			return er
		}
		d, er := losslessxml.Parse(raw)
		if er != nil {
			return er
		}
		if len(d.Elements()) == 0 || d.Elements()[0].Name() != name(ns, local) {
			return graphicsSmartArtUnsupported("dependency root")
		}
		return nil
	}
	edgeRecord := func(owner string, r packaging.Edge) SmartArtEdge {
		var target *string
		if !r.External && r.ResolvedPart != "" {
			v := r.ResolvedPart
			target = &v
		}
		return SmartArtEdge{owner, r.ID, r.Type, r.Target, r.External, target}
	}
	frames := []losslessxml.Element{}
	var visit func(losslessxml.Element)
	visit = func(parent losslessxml.Element) {
		for _, n := range graphicsChildren(doc, parent) {
			if n.Name().Space != packaging.NSPresentationML {
				continue
			}
			switch n.Name().Local {
			case "grpSp":
				visit(n)
			case "graphicFrame":
				frames = append(frames, n)
			}
		}
	}
	visit(tree)
	result := []SmartArtInfo{}
	used := map[losslessxml.Element]bool{}
	for _, frame := range frames {
		matches := false
		for _, graphic := range graphicsChildren(doc, frame) {
			if graphic.Name() != name(packaging.NSDrawingML, "graphic") {
				continue
			}
			for _, data := range graphicsChildren(doc, graphic) {
				uri, _ := graphicsAttr(data, "", "uri")
				if data.Name() == name(packaging.NSDrawingML, "graphicData") && uri == graphicsDiagramNS {
					matches = true
				}
			}
		}
		if !matches {
			continue
		}
		graphic, er := graphicsSmartOne(doc, frame, packaging.NSDrawingML, "graphic")
		if er != nil {
			return nil, er
		}
		data, er := graphicsSmartOne(doc, graphic, packaging.NSDrawingML, "graphicData")
		if er != nil {
			return nil, er
		}
		leaf, er := graphicsSmartOne(doc, data, graphicsDiagramNS, "relIds")
		if er != nil {
			return nil, er
		}
		text, _ := leaf.Text()
		if len(graphicsChildren(doc, data)) != 1 || len(graphicsChildren(doc, leaf)) != 0 || strings.TrimSpace(text) != "" {
			return nil, graphicsSmartArtUnsupported("ambiguous role leaf")
		}
		used[leaf] = true
		for _, a := range leaf.Attributes() {
			if (a.Name.Local == "dm" || a.Name.Local == "lo" || a.Name.Local == "qs" || a.Name.Local == "cs") && a.Name.Space != packaging.NSDocumentRelationships {
				return nil, graphicsSmartArtUnsupported("foreign role attribute")
			}
		}
		nv, er := graphicsSmartOne(doc, frame, packaging.NSPresentationML, "nvGraphicFramePr")
		if er != nil {
			return nil, er
		}
		identity, er := graphicsSmartOne(doc, nv, packaging.NSPresentationML, "cNvPr")
		if er != nil {
			return nil, er
		}
		idRaw, _ := graphicsAttr(identity, "", "id")
		id, _ := graphicsNumber(idRaw, 1, math.MaxInt32)
		label, _ := graphicsAttr(identity, "", "name")
		info := SmartArtInfo{ShapeID: uint32(id), Name: label, SlidePart: part, Roots: []SmartArtRoot{}, Parts: []SmartArtPart{}, Edges: []SmartArtEdge{}, DrawingParts: []string{}, Limits: SmartArtLimits{Reason: "Dependency inspection only; no SmartArt layout/edit engine."}}
		for _, role := range [][5]string{{"data", "dm", "diagramData", "diagramData", "dataModel"}, {"layout", "lo", "diagramLayout", "diagramLayout", "layoutDef"}, {"style", "qs", "diagramQuickStyle", "diagramStyle", "styleDef"}, {"colors", "cs", "diagramColors", "diagramColors", "colorsDef"}} {
			rid, ok := graphicsAttr(leaf, packaging.NSDocumentRelationships, role[1])
			count := 0
			var rel packaging.Edge
			for _, edge := range byOwner[part] {
				if edge.ID == rid {
					count++
					rel = edge
				}
			}
			if !ok || rid == "" || count != 1 || rel.Type != graphicsRelationNS+"/"+role[2] || rel.External || rel.ResolvedPart == "" {
				return nil, graphicsSmartArtUnsupported("role relationship")
			}
			if er = root(rel.ResolvedPart, "application/vnd.openxmlformats-officedocument.drawingml."+role[3]+"+xml", role[4], graphicsDiagramNS); er != nil {
				return nil, er
			}
			a, er := asset(rel.ResolvedPart)
			if er != nil {
				return nil, er
			}
			info.Roots = append(info.Roots, SmartArtRoot{a, role[0], rid, rel.Target})
		}
		parts := map[string]SmartArtPart{}
		drawings := map[string]bool{}
		pending := []string{}
		for i := len(info.Roots) - 1; i >= 0; i-- {
			pending = append(pending, info.Roots[i].PartName)
		}
		for len(pending) > 0 {
			owner := pending[len(pending)-1]
			pending = pending[:len(pending)-1]
			if _, ok := parts[owner]; ok {
				continue
			}
			if len(parts) >= 256 {
				return nil, graphicsSmartArtUnsupported("part limit")
			}
			a, er := asset(owner)
			if er != nil {
				return nil, er
			}
			parts[owner] = a
			deps := byOwner[owner]
			if types[owner] == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" {
				raw, _, er := s.pkg.Part(owner)
				if er != nil {
					return nil, er
				}
				d, er := losslessxml.Parse(raw)
				if er != nil {
					return nil, er
				}
				metadata := []losslessxml.Element{}
				for _, n := range d.Elements() {
					if n.Name() == name(graphicsPersistDiagramNS, "dataModelExt") {
						metadata = append(metadata, n)
					}
				}
				if len(metadata) > 1 {
					return nil, graphicsSmartArtUnsupported("ambiguous drawing metadata")
				}
				for _, m := range metadata {
					rid, _ := graphicsAttr(m, "", "relId")
					local, slide := []packaging.Edge{}, []packaging.Edge{}
					for _, edge := range deps {
						if edge.ID == rid && edge.Type == graphicsDiagramDrawingRel && !edge.External {
							local = append(local, edge)
						}
					}
					for _, edge := range byOwner[part] {
						if edge.ID == rid && edge.Type == graphicsDiagramDrawingRel && !edge.External {
							slide = append(slide, edge)
						}
					}
					if rid == "" || len(local)+len(slide) != 1 {
						return nil, graphicsSmartArtUnsupported("stale/ambiguous drawing metadata")
					}
					if len(slide) == 1 {
						r := slide[0]
						if er = root(r.ResolvedPart, "application/vnd.ms-office.drawingml.diagramDrawing+xml", "drawing", graphicsPersistDiagramNS); er != nil {
							return nil, er
						}
						drawings[r.ResolvedPart] = true
						seen := false
						for _, edge := range info.Edges {
							if edge.Owner == part && edge.RelationshipID == r.ID {
								seen = true
							}
						}
						if !seen {
							if len(info.Edges) >= 1024 {
								return nil, graphicsSmartArtUnsupported("edge limit")
							}
							info.Edges = append(info.Edges, edgeRecord(part, r))
						}
						pending = append(pending, r.ResolvedPart)
					}
				}
			}
			for _, r := range deps {
				if len(info.Edges) >= 1024 {
					return nil, graphicsSmartArtUnsupported("edge limit")
				}
				info.Edges = append(info.Edges, edgeRecord(owner, r))
				if r.External {
					continue
				}
				if r.ResolvedPart == "" {
					return nil, graphicsSmartArtUnsupported("unresolved dependency")
				}
				if r.Type == graphicsDiagramDrawingRel {
					if er = root(r.ResolvedPart, "application/vnd.ms-office.drawingml.diagramDrawing+xml", "drawing", graphicsPersistDiagramNS); er != nil {
						return nil, er
					}
					drawings[r.ResolvedPart] = true
				}
				pending = append(pending, r.ResolvedPart)
			}
		}
		for _, p := range parts {
			info.Parts = append(info.Parts, p)
		}
		sort.Slice(info.Parts, func(i, j int) bool { return info.Parts[i].PartName < info.Parts[j].PartName })
		sort.Slice(info.Edges, func(i, j int) bool {
			if info.Edges[i].Owner == info.Edges[j].Owner {
				return info.Edges[i].RelationshipID < info.Edges[j].RelationshipID
			}
			return info.Edges[i].Owner < info.Edges[j].Owner
		})
		for p := range drawings {
			info.DrawingParts = append(info.DrawingParts, p)
		}
		sort.Strings(info.DrawingParts)
		result = append(result, info)
	}
	for _, n := range es {
		if n.Name() == name(graphicsDiagramNS, "relIds") && !used[n] {
			return nil, graphicsSmartArtUnsupported("misplaced role leaf")
		}
	}
	return result, nil
}
