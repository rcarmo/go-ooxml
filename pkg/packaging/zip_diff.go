package packaging

import (
	"bytes"
	"sort"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// DiffZIP32 compares two raw ZIP32 archives without requiring an OPC registry.
// Both archives are admitted independently before any report is returned.
// Payload equality precedes bounded conservative XML equivalence; suffixes
// select XML candidates, not arbitrary XML-looking binary payloads.
func DiffZIP32(left, right []byte, limits ZIP32Limits) (PackageDiff, error) {
	a, err := ReadZIP32(left, limits)
	if err != nil {
		return PackageDiff{}, err
	}
	b, err := ReadZIP32(right, limits)
	if err != nil {
		return PackageDiff{}, err
	}
	result := PackageDiff{Added: []string{}, Removed: []string{}, Changed: []string{}, EquivalentXML: []string{}}
	old := make(map[string][]byte, len(a))
	newParts := make(map[string][]byte, len(b))
	for _, entry := range a {
		old[entry.Name] = entry.Data
	}
	for _, entry := range b {
		newParts[entry.Name] = entry.Data
	}
	for name, before := range old {
		after, exists := newParts[name]
		if !exists {
			result.Removed = append(result.Removed, name)
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
	for name := range newParts {
		if _, exists := old[name]; !exists {
			result.Added = append(result.Added, name)
		}
	}
	sort.Strings(result.Added)
	sort.Strings(result.Removed)
	sort.Strings(result.Changed)
	sort.Strings(result.EquivalentXML)
	return result, nil
}
