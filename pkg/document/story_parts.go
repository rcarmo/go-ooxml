package document

import (
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"sort"
)

func (s *EditSession) storyParts() ([]string, error) {
	g, err := s.pkg.Graph()
	if err != nil {
		return nil, err
	}
	kinds := map[string]bool{packaging.RelTypeHeader: true, packaging.RelTypeFooter: true, packaging.RelTypeFootnotes: true, packaging.RelTypeEndnotes: true, packaging.RelTypeComments: true}
	parts := []string{s.part}
	seen := map[string]bool{s.part: true}
	for i := 0; i < len(parts); i++ {
		for _, e := range g.Edges {
			if e.Source == parts[i] && kinds[e.Type] && !e.External && !seen[e.ResolvedPart] {
				parts = append(parts, e.ResolvedPart)
				seen[e.ResolvedPart] = true
			}
		}
	}
	sort.Strings(parts[1:])
	return parts, nil
}
