package presentation

import (
	"bytes"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type PictureDeleteOptions struct {
	CollectMedia bool `json:"collectMedia"`
}
type PictureDeleteReceipt struct {
	ShapeID              uint32   `json:"shapeId"`
	PartName             string   `json:"partName"`
	RemovedRelationships []string `json:"removedRelationships"`
	RemovedMedia         []string `json:"removedMedia"`
}

var graphicsCollectableMedia = regexp.MustCompile(`(?i)^ppt/media/[^/]+\.(?:png|jpe?g|svg)$`)
var graphicsXMLDeclaration = regexp.MustCompile(`^\x{FEFF}?\s*<\?xml\s[^?]*\?>`)

// DeletePicture removes exactly one picture; collection needs package-local absence proofs.
func (s *EditSession) DeletePicture(part string, id uint32, options PictureDeleteOptions) (PictureDeleteReceipt, error) {
	var none PictureDeleteReceipt
	if id == 0 || id > math.MaxInt32 {
		return none, graphicsUnsupported("bounded delete identity")
	}
	if e := s.manipulationProtection(); e != nil {
		return none, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	_, doc, picture, source, hash, e := s.graphicsPictureSelection(part, id)
	if e != nil {
		return none, e
	}
	for _, n := range doc.Elements() {
		if n.Name() != name(packaging.NSDrawingML, "stCxn") && n.Name() != name(packaging.NSDrawingML, "endCxn") {
			continue
		}
		raw, _ := graphicsAttr(n, "", "id")
		value, _ := strconv.ParseUint(raw, 10, 32)
		if uint32(value) == id {
			return none, graphicsUnsupported("attached connector endpoint")
		}
	}
	graph, e := s.pkg.Graph()
	if e != nil {
		return none, e
	}
	dependencies := []packaging.Edge{}
	seen := map[string]bool{}
	for _, n := range doc.Elements() {
		if n != picture && !manipulationWithin(n, picture) {
			continue
		}
		for _, local := range []string{"embed", "link", "id"} {
			value, ok := graphicsAttr(n, packaging.NSDocumentRelationships, local)
			if !ok {
				continue
			}
			count := 0
			var edge packaging.Edge
			for _, candidate := range graph.Edges {
				if candidate.Source == part && candidate.ID == value {
					count++
					edge = candidate
				}
			}
			if count != 1 || local != "id" && edge.Type != packaging.RelTypeImage {
				return none, graphicsUnsupported("ambiguous picture dependency")
			}
			if edge.Type == packaging.RelTypeImage && !seen[edge.ID] {
				dependencies = append(dependencies, edge)
				seen[edge.ID] = true
			}
		}
	}
	a, z := picture.SourceRange()
	next := append(append([]byte{}, source[:a]...), source[z:]...)
	remaining, e := losslessxml.Parse(next)
	if e != nil {
		return none, e
	}
	lexical := graphicsXMLDeclaration.ReplaceAll(next, nil)
	canCollect := options.CollectMedia && !bytes.Contains(lexical, []byte("<!--")) && !bytes.Contains(lexical, []byte("<?")) && !bytes.Contains(lexical, []byte("<![CDATA["))
	for _, n := range remaining.Elements() {
		text, _ := n.Text()
		if strings.TrimSpace(text) != "" && n.Name() != name(packaging.NSDrawingML, "t") {
			canCollect = false
		}
	}
	receipt := PictureDeleteReceipt{id, part, []string{}, []string{}}
	change := packaging.GraphMutation{Replacements: []packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: next}}}
	removed := map[string]bool{}
	names := map[string]bool{}
	types := map[string]string{}
	for _, p := range graph.Parts {
		names[p.Name] = true
		types[p.Name] = p.ContentType
	}
	if canCollect {
		for _, edge := range dependencies {
			used := false
			for _, n := range remaining.Elements() {
				for _, attr := range n.Attributes() {
					if attr.Value == edge.ID {
						used = true
					}
				}
				for _, v := range n.Namespaces() {
					if v == edge.ID {
						used = true
					}
				}
			}
			if used {
				continue
			}
			change.Removals = append(change.Removals, packaging.RelationshipRemoval{Source: part, ID: edge.ID})
			receipt.RemovedRelationships = append(receipt.RemovedRelationships, edge.ID)
			removed[edge.ID] = true
		}
		deleted := map[string]bool{}
		for _, edge := range dependencies {
			target := edge.ResolvedPart
			if edge.External || target == "" || deleted[target] || !removed[edge.ID] || !graphicsCollectableMedia.MatchString(target) || !strings.HasPrefix(types[target], "image/") || names[packaging.RelationshipsPathForPart(target)] {
				continue
			}
			incoming := false
			for _, candidate := range graph.Edges {
				if !candidate.External && candidate.ResolvedPart == target && !(candidate.Source == part && removed[candidate.ID]) {
					incoming = true
					break
				}
			}
			if incoming {
				continue
			}
			_, h, er := s.pkg.Part(target)
			if er != nil {
				return none, er
			}
			change.Deletions = append(change.Deletions, packaging.PartDeletion{Name: target, ExpectedSHA256: h})
			receipt.RemovedMedia = append(receipt.RemovedMedia, target)
			deleted[target] = true
		}
	}
	plan, e := s.pkg.PlanGraphMutation(change)
	if e != nil {
		return none, e
	}
	if e = s.pkg.ApplyGraphPlan(plan); e != nil {
		return none, e
	}
	s.generation++
	return receipt, nil
}
