package packaging

import (
	"bytes"
	"sort"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// PackageDiff compares two independent immutable package snapshots. Lists are
// disjoint, sorted, and detached. A changed effective MIME is a changed part
// even if its payload bytes are identical. This is not a rendering comparison.
type PackageDiff struct {
	Added         []string `json:"added"`
	Removed       []string `json:"removed"`
	Changed       []string `json:"changed"`
	EquivalentXML []string `json:"equivalent_xml"`
}

// DiffPreserved compares current payloads and effective content types. XML
// equivalence is conservative and bounded by losslessxml.Equivalent. Distinct
// inputs and staged sessions are not mutated or reserialized by this operation.
func DiffPreserved(left, right *Preserved) (PackageDiff, error) {
	result := PackageDiff{Added: []string{}, Removed: []string{}, Changed: []string{}, EquivalentXML: []string{}}
	if left == nil || right == nil {
		return result, graphEditError("missing_target", "", "nil package snapshot")
	}
	leftGraph, err := left.Graph()
	if err != nil {
		return result, err
	}
	rightGraph, err := right.Graph()
	if err != nil {
		return result, err
	}
	types := func(graph Graph) map[string]string {
		out := map[string]string{}
		for _, part := range graph.Parts {
			out[part.Name] = part.ContentType
		}
		return out
	}
	lt, rt := types(leftGraph), types(rightGraph)
	// [Content_Types].xml is the registry controlling effective MIME but is
	// not itself in Graph.Parts. Compare its actual payload independently.
	leftRegistry, _, err := left.Part(ContentTypesPath)
	if err != nil {
		return result, err
	}
	rightRegistry, _, err := right.Part(ContentTypesPath)
	if err != nil {
		return result, err
	}
	if !bytes.Equal(leftRegistry, rightRegistry) {
		result.Changed = append(result.Changed, ContentTypesPath)
	}
	seen := map[string]bool{}
	for name, oldType := range lt {
		seen[name] = true
		newType, ok := rt[name]
		if !ok {
			result.Removed = append(result.Removed, name)
			continue
		}
		before, _, err := left.Part(name)
		if err != nil {
			return PackageDiff{}, err
		}
		after, _, err := right.Part(name)
		if err != nil {
			return PackageDiff{}, err
		}
		if oldType != newType {
			result.Changed = append(result.Changed, name)
			continue
		}
		if bytes.Equal(before, after) {
			continue
		}
		lower := strings.ToLower(name)
		if (strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".rels")) && losslessxml.Equivalent(before, after) {
			result.EquivalentXML = append(result.EquivalentXML, name)
		} else {
			result.Changed = append(result.Changed, name)
		}
	}
	for name := range rt {
		if !seen[name] {
			result.Added = append(result.Added, name)
		}
	}
	sort.Strings(result.Added)
	sort.Strings(result.Removed)
	sort.Strings(result.Changed)
	sort.Strings(result.EquivalentXML)
	return result, nil
}
