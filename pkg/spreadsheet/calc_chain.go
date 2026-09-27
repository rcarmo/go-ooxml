package spreadsheet

import (
	"strings"

	"github.com/rcarmo/go-ooxml/internal/formula"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type calculationChain struct{ part, id, hash string }

// calculationChain proves ownership of the optional cache/order part. Its
// entries are not used as dependency evidence; the formula analyser supplies it.
func (s *EditSession) calculationChain(g packaging.Graph) (calculationChain, error) {
	chain := calculationChain{}
	parts := map[string]packaging.GraphPart{}
	for _, p := range g.Parts {
		parts[p.Name] = p
	}
	for _, edge := range g.Edges {
		if edge.Type != packaging.RelTypeCalcChain {
			continue
		}
		if chain.part != "" || edge.Source != s.main || edge.External || strings.Contains(edge.Target, "#") {
			return chain, editRefusal("unsupported_structure", "ambiguous or external calculation chain")
		}
		chain.part, chain.id = edge.ResolvedPart, edge.ID
	}
	for _, p := range g.Parts {
		if p.ContentType == packaging.ContentTypeCalcChain || strings.EqualFold(p.Name, "xl/calcChain.xml") {
			if p.Name != chain.part || p.ContentType != packaging.ContentTypeCalcChain {
				return chain, editRefusal("unsupported_structure", "unowned or mistyped calculation chain")
			}
		}
	}
	if chain.part == "" {
		return chain, nil
	}
	if parts[chain.part].ContentType != packaging.ContentTypeCalcChain || parts[chain.part].Inbound != 1 {
		return chain, editRefusal("unsupported_structure", "shared or mistyped calculation chain")
	}
	if _, _, err := s.pkg.Part(packaging.RelationshipsPathForPart(chain.part)); err == nil {
		return chain, editRefusal("unsupported_structure", "calculation chain has outgoing registry")
	}
	data, hash, err := s.pkg.Part(chain.part)
	if err != nil {
		return chain, err
	}
	chain.hash = hash
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return chain, editRefusal("unsupported_structure", "invalid calculation chain XML")
	}
	es := doc.Elements()
	root := es[0]
	if root.Name() != expanded("calcChain") || len(root.Attributes()) != 0 {
		return chain, editRefusal("unsupported_structure", "unknown calculation chain root")
	}
	if text, _ := root.Text(); strings.TrimSpace(text) != "" {
		return chain, editRefusal("unsupported_structure", "mixed chain content")
	}
	for _, e := range es[1:] {
		parent, ok := e.Parent()
		text, leaf := e.Text()
		if !ok || parent != root || e.Name() != expanded("c") || !leaf || strings.TrimSpace(text) != "" {
			return chain, editRefusal("unsupported_structure", "extended calculation chain")
		}
		r := attr(e, "r")
		if _, err = formula.ParseCell(r); err != nil || !strictCell.MatchString(r) {
			return chain, editRefusal("unsupported_structure", "noncanonical calculation chain cell")
		}
		for _, a := range e.Attributes() {
			if a.Name.Space != "" {
				return chain, editRefusal("unsupported_structure", "extended chain attribute")
			}
			switch a.Name.Local {
			case "r":
			case "i":
				n, err := unsignedIndex(a.Value)
				if err != nil || n == 0 {
					return chain, editRefusal("unsupported_structure", "invalid chain sheet ID")
				}
			case "l", "s", "a", "t":
				if a.Value != "0" && a.Value != "1" && a.Value != "true" && a.Value != "false" {
					return chain, editRefusal("unsupported_structure", "invalid chain boolean")
				}
			default:
				return chain, editRefusal("unsupported_structure", "unknown chain attribute")
			}
		}
	}
	return chain, nil
}
