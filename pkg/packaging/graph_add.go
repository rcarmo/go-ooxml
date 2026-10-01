package packaging

import (
	"encoding/xml"
	"mime"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// planGraphAdditions alters only the private plan. Final Graph() validation in
// PlanGraphMutation checks all resulting ownership and target invariants.
func (p *Preserved) planGraphAdditions(change GraphMutation, graph Graph, plan *GraphPlan) error {
	current := func(name string) ([]byte, error) {
		if b, ok := plan.patches[name]; ok {
			return b, nil
		}
		b, _, err := p.Part(name)
		return b, err
	}
	if len(change.Relationships) > 0 {
		type key struct{ source, id string }
		seen := map[key]bool{}
		for _, edge := range graph.Edges {
			seen[key{edge.Source, edge.ID}] = true
		}
		groups := map[string][]losslessxml.NewElement{}
		for _, addition := range change.Relationships {
			if addition.ID == "" || addition.Type == "" || strings.TrimSpace(addition.ID) != addition.ID || strings.TrimSpace(addition.Type) != addition.Type || strings.ContainsAny(addition.ID, " \t\n\r") {
				return graphEditError("relationship_policy", addition.Source, "invalid relationship ID or type")
			}
			identity := key{addition.Source, addition.ID}
			if seen[identity] {
				return graphEditError("ambiguous_target", addition.Source, "duplicate relationship ID")
			}
			seen[identity] = true
			if addition.Source != "" && !p.hasPart(addition.Source) {
				if _, pending := plan.additions[addition.Source]; !pending {
					return graphEditError("missing_target", addition.Source, "relationship owner absent")
				}
			}
			if err := graphPartName(addition.TargetPart); err != nil {
				return err
			}
			if !p.hasPart(addition.TargetPart) {
				if _, ok := plan.additions[addition.TargetPart]; !ok {
					return graphEditError("missing_target", addition.TargetPart, "relationship target absent")
				}
			}
			registryName := PackageRelsPath
			if addition.Source != "" {
				registryName = RelationshipsPathForPart(addition.Source)
			}
			groups[registryName] = append(groups[registryName], losslessxml.NewElement{Name: xml.Name{Space: relNS, Local: "Relationship"}, Attributes: []xml.Attr{{Name: xml.Name{Local: "Id"}, Value: addition.ID}, {Name: xml.Name{Local: "Type"}, Value: addition.Type}, {Name: xml.Name{Local: "Target"}, Value: retargetURI(addition.Source, addition.TargetPart, "")}}})
		}
		for name, nodes := range groups {
			var data []byte
			if p.hasPart(name) {
				var err error
				data, err = current(name)
				if err != nil {
					return err
				}
			} else {
				data = []byte(`<Relationships xmlns="` + relNS + `"></Relationships>`)
			}
			doc, err := losslessxml.Parse(data)
			if err != nil {
				return graphEditError("relationship_policy", name, err.Error())
			}
			elements := doc.Elements()
			if len(elements) == 0 || elements[0].Name() != (xml.Name{Space: relNS, Local: "Relationships"}) {
				return graphEditError("relationship_policy", name, "unexpected registry root")
			}
			updated, err := doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: elements[0], Children: nodes}})
			if err != nil {
				return graphEditError("relationship_policy", name, err.Error())
			}
			if p.hasPart(name) {
				plan.patches[name] = updated
			} else {
				plan.additions[name] = updated
			}
		}
	}
	if len(change.ContentTypes) > 0 {
		data, err := current(ContentTypesPath)
		if err != nil {
			return err
		}
		doc, err := losslessxml.Parse(data)
		if err != nil {
			return graphEditError("relationship_policy", ContentTypesPath, err.Error())
		}
		roots := doc.Elements()
		if len(roots) == 0 || roots[0].Name() != (xml.Name{Space: ctNS, Local: "Types"}) {
			return graphEditError("relationship_policy", ContentTypesPath, "unexpected content-type registry")
		}
		existing := map[string]losslessxml.Element{}
		for _, entry := range roots[1:] {
			if entry.Name() != (xml.Name{Space: ctNS, Local: "Override"}) {
				continue
			}
			attrs, err := attributes(entry, "PartName", "ContentType")
			if err != nil {
				return graphEditError("relationship_policy", ContentTypesPath, err.Error())
			}
			part, err := resolveGraphTarget("", attrs["PartName"])
			if err != nil {
				return graphEditError("relationship_policy", ContentTypesPath, err.Error())
			}
			existing[part] = entry
		}
		seen := map[string]bool{}
		edits := []losslessxml.AttributeEdit{}
		inserts := []losslessxml.NewElement{}
		for _, change := range change.ContentTypes {
			if err := graphPartName(change.Part); err != nil {
				return err
			}
			if seen[change.Part] {
				return graphEditError("ambiguous_target", change.Part, "duplicate content-type edit")
			}
			seen[change.Part] = true
			if !p.hasPart(change.Part) {
				return graphEditError("missing_target", change.Part, "content-type target absent")
			}
			typ, params, e := mime.ParseMediaType(change.ContentType)
			if e != nil || len(params) != 0 || typ != change.ContentType || !strings.Contains(typ, "/") {
				return graphEditError("relationship_policy", change.Part, "canonical MIME type required")
			}
			if existingEntry, ok := existing[change.Part]; ok {
				edits = append(edits, losslessxml.AttributeEdit{Target: existingEntry, Name: xml.Name{Local: "ContentType"}, Value: typ})
			} else {
				inserts = append(inserts, losslessxml.NewElement{Name: xml.Name{Space: ctNS, Local: "Override"}, Attributes: []xml.Attr{{Name: xml.Name{Local: "PartName"}, Value: partURI(change.Part)}, {Name: xml.Name{Local: "ContentType"}, Value: typ}}})
			}
		}
		if len(edits) > 0 {
			data, err = doc.Edit(nil, edits)
			if err != nil {
				return graphEditError("relationship_policy", ContentTypesPath, err.Error())
			}
		}
		if len(inserts) > 0 {
			doc, err = losslessxml.Parse(data)
			if err != nil {
				return err
			}
			data, err = doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: doc.Elements()[0], Children: inserts}})
			if err != nil {
				return graphEditError("relationship_policy", ContentTypesPath, err.Error())
			}
		}
		plan.patches[ContentTypesPath] = data
	}
	return nil
}
