package packaging

import (
	"bytes"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// planGraphRemovals extends a private plan only. Registry edits may compose with
// additions/retargets, but edits selecting the same edge or part conflict.
func (p *Preserved) planGraphRemovals(change GraphMutation, plan *GraphPlan) error {
	current := func(name string) ([]byte, error) {
		if b, ok := plan.patches[name]; ok {
			return b, nil
		}
		b, _, err := p.Part(name)
		return b, err
	}
	for _, d := range change.Deletions {
		if err := graphPartName(d.Name); err != nil {
			return err
		}
		if plan.deletions[d.Name] {
			return graphEditError("ambiguous_target", d.Name, "duplicate deletion")
		}
		if _, ok := plan.additions[d.Name]; ok {
			return graphEditError("ambiguous_target", d.Name, "add/delete conflict")
		}
		_, hash, err := p.Part(d.Name)
		if err != nil {
			return err
		}
		if d.ExpectedSHA256 == "" || hash != d.ExpectedSHA256 {
			return graphEditError("stale_target", d.Name, "deletion fingerprint changed")
		}
		if p.hasPart(RelationshipsPathForPart(d.Name)) {
			return graphEditError("relationship_policy", d.Name, "deletion requires a leaf without a relationship registry")
		}
		plan.deletions[d.Name] = true
	}
	type edgeKey struct{ source, id string }
	seen := map[edgeKey]bool{}
	for _, r := range change.Retargets {
		seen[edgeKey{r.Source, r.ID}] = true
	}
	grouped := map[string][]string{}
	for _, r := range change.Removals {
		key := edgeKey{r.Source, r.ID}
		if seen[key] {
			return graphEditError("ambiguous_target", r.Source, "duplicate/conflicting relationship removal")
		}
		seen[key] = true
		name := PackageRelsPath
		if r.Source != "" {
			if !p.hasPart(r.Source) {
				return graphEditError("missing_target", r.Source, "relationship owner absent")
			}
			name = RelationshipsPathForPart(r.Source)
		}
		grouped[name] = append(grouped[name], r.ID)
	}
	for name, ids := range grouped {
		data, err := current(name)
		if err != nil {
			return err
		}
		doc, err := losslessxml.Parse(data)
		if err != nil {
			return graphEditError("relationship_policy", name, err.Error())
		}
		entries := map[string]losslessxml.Element{}
		for _, e := range doc.Elements()[1:] {
			attrs, err := attributes(e, "Id", "Type", "Target", "TargetMode")
			if err != nil {
				return graphEditError("relationship_policy", name, err.Error())
			}
			entries[attrs["Id"]] = e
		}
		targets := []losslessxml.Element{}
		for _, id := range ids {
			e, ok := entries[id]
			if !ok {
				return graphEditError("missing_target", name, "removed edge absent")
			}
			targets = append(targets, e)
		}
		data, err = doc.RemoveElements(targets)
		if err != nil {
			return graphEditError("relationship_policy", name, err.Error())
		}
		plan.patches[name] = data
	}
	if len(plan.deletions) > 0 {
		data, err := current(ContentTypesPath)
		if err != nil {
			return err
		}
		doc, err := losslessxml.Parse(data)
		if err != nil {
			return err
		}
		targets := []losslessxml.Element{}
		for _, e := range doc.Elements()[1:] {
			if e.Name().Local != "Override" {
				continue
			}
			a, err := attributes(e, "PartName", "ContentType")
			if err != nil {
				return err
			}
			part, err := resolveGraphTarget("", a["PartName"])
			if err != nil {
				return err
			}
			if plan.deletions[part] {
				targets = append(targets, e)
			}
		}
		if len(targets) > 0 {
			data, err = doc.RemoveElements(targets)
			if err != nil {
				return graphEditError("relationship_policy", ContentTypesPath, err.Error())
			}
			plan.patches[ContentTypesPath] = data
		}
	}
	replaced := map[string]bool{}
	for _, r := range change.Replacements {
		if err := graphPartName(r.Part); err != nil {
			return err
		}
		if replaced[r.Part] || plan.deletions[r.Part] {
			return graphEditError("ambiguous_target", r.Part, "duplicate/conflicting payload edit")
		}
		replaced[r.Part] = true
		if _, ok := plan.additions[r.Part]; ok {
			return graphEditError("ambiguous_target", r.Part, "add/replace conflict")
		}
		before, hash, err := p.Part(r.Part)
		if err != nil {
			return err
		}
		if hash != r.ExpectedSHA256 || r.ExpectedSHA256 == "" {
			return graphEditError("stale_target", r.Part, "replacement fingerprint changed")
		}
		if strings.HasSuffix(strings.ToLower(r.Part), ".xml") {
			if err = validateXMLPayload(r.Data); err != nil {
				return invalidPart("graph_edit", r.Part, err.Error())
			}
		}
		if !bytes.Equal(before, r.Data) {
			plan.patches[r.Part] = bytes.Clone(r.Data)
		}
	}
	// Final Graph() rejects every dangling inbound edge after selected removals.
	return nil
}
