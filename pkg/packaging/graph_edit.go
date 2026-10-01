package packaging

import (
	"bytes"
	"encoding/xml"
	"mime"
	"net/url"
	"path"
	"sort"
	"strings"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// PartAddition adds a new payload with an explicit content-type override.
type PartAddition struct {
	Name, ContentType string
	Data              []byte
}

// RelationshipRetarget changes one existing internal edge to a canonical part name.
// Source is empty for the package root; TargetPart is never a URI or external URL.
type RelationshipRetarget struct{ Source, ID, TargetPart string }

// RelationshipAddition creates a new internal edge to an existing or jointly
// added part. Source is empty for the package root.
type RelationshipAddition struct{ Source, ID, Type, TargetPart string }

// ContentTypeChange changes the effective MIME of one existing part by a
// scoped Override, leaving the payload unchanged.
type ContentTypeChange struct{ Part, ContentType string }

// PartDeletion requires a current payload fingerprint; only detached leaf parts may be removed.
type PartDeletion struct{ Name, ExpectedSHA256 string }

// RelationshipRemoval removes a selected edge; format adapters must prove its owner has no live reference.
type RelationshipRemoval struct{ Source, ID string }
type GraphMutation struct {
	Relationships []RelationshipAddition
	ContentTypes  []ContentTypeChange
	Deletions     []PartDeletion
	Removals      []RelationshipRemoval
	Replacements  []Replacement
	Additions     []PartAddition
	Retargets     []RelationshipRetarget
}

// GraphPlan owns private payloads and is bound to a single session generation.
// It proves OPC registry consistency, not application-level dependency semantics.
type GraphPlan struct {
	owner              *Preserved
	generation         uint64
	additions, patches map[string][]byte
	deletions          map[string]bool
	consumed           bool
}

func graphEditError(kind, part, detail string) error {
	return &Refusal{Kind: kind, Operation: "graph_edit", Part: part, Detail: detail}
}
func reservedGraphPart(name string) bool {
	lower := strings.ToLower(name)
	return lower == strings.ToLower(ContentTypesPath) || strings.HasSuffix(lower, ".rels") || strings.HasPrefix(lower, "_xmlsignatures/") || strings.HasPrefix(lower, "_rels/") || strings.Contains(lower, "/_rels/")
}
func graphPartName(name string) error {
	if err := validateMemberName(name, false); err != nil {
		return err
	}
	if !utf8.ValidString(name) || strings.Contains(name, "%") || strings.HasSuffix(name, "/") {
		return graphEditError("relationship_policy", name, "noncanonical or escaped part name")
	}
	for _, segment := range strings.Split(name, "/") {
		if strings.HasSuffix(segment, ".") {
			return graphEditError("relationship_policy", name, "trailing-dot part segment")
		}
	}
	if reservedGraphPart(name) {
		return graphEditError("relationship_policy", name, "registry/signature parts require another policy")
	}
	return nil
}
func partURI(name string) string { return (&url.URL{Path: "/" + name}).EscapedPath() }

// Keep the original absolute/relative form, with POSIX package paths on all OSes.
func retargetURI(source, target, original string) string {
	if strings.HasPrefix(original, "/") {
		return partURI(target)
	}
	var from []string
	dir := path.Dir(source)
	if dir != "." {
		from = strings.Split(dir, "/")
	}
	to := strings.Split(target, "/")
	same := 0
	for same < len(from) && same < len(to) && from[same] == to[same] {
		same++
	}
	segments := make([]string, 0, len(from)-same+len(to)-same)
	for i := same; i < len(from); i++ {
		segments = append(segments, "..")
	}
	segments = append(segments, to[same:]...)
	return (&url.URL{Path: strings.Join(segments, "/")}).EscapedPath()
}

// PlanGraphMutation prepares all new parts, type overrides and existing internal
// edge retargets without modifying the session. Original shared targets remain.
// Unknown registries, collisions, signed input, external edges and fragments
// refuse. Callers must prove format-specific ownership before applying the plan.
func (p *Preserved) PlanGraphMutation(change GraphMutation) (*GraphPlan, error) {
	plan := &GraphPlan{owner: p, generation: p.generation, additions: map[string][]byte{}, patches: map[string][]byte{}, deletions: map[string]bool{}}
	if len(change.Additions) == 0 && len(change.Retargets) == 0 && len(change.Deletions) == 0 && len(change.Removals) == 0 && len(change.Replacements) == 0 && len(change.Relationships) == 0 && len(change.ContentTypes) == 0 {
		return plan, nil
	}
	if p.signed {
		return nil, graphEditError("unsupported_structure", "", "signed packages require an explicit policy")
	}
	graph, err := p.Graph()
	if err != nil {
		return nil, err
	}
	names := map[string]bool{}
	for _, name := range p.partNames() {
		names[strings.ToLower(name)] = true
	}
	// Deleted original identities cannot be reused by additions in this API.
	for name := range p.parts {
		names[strings.ToLower(name)] = true
	}
	additions := append([]PartAddition(nil), change.Additions...)
	sort.Slice(additions, func(i, j int) bool { return additions[i].Name < additions[j].Name })
	nodes := []losslessxml.NewElement{}
	for _, a := range additions {
		if err = graphPartName(a.Name); err != nil {
			return nil, err
		}
		fold := strings.ToLower(a.Name)
		if names[fold] {
			return nil, graphEditError("ambiguous_target", a.Name, "existing or case-colliding part")
		}
		names[fold] = true
		typ, params, e := mime.ParseMediaType(a.ContentType)
		if e != nil || len(params) != 0 || !strings.Contains(typ, "/") || a.ContentType != typ {
			return nil, graphEditError("relationship_policy", a.Name, "canonical parameter-free MIME type required")
		}
		data := bytes.Clone(a.Data)
		if strings.HasSuffix(strings.ToLower(a.Name), ".xml") || strings.HasSuffix(typ, "+xml") || typ == "application/xml" || typ == "text/xml" {
			if e = validateXMLPayload(data); e != nil {
				return nil, invalidPart("graph_edit", a.Name, e.Error())
			}
		}
		plan.additions[a.Name] = data
		nodes = append(nodes, losslessxml.NewElement{Name: xml.Name{Space: ctNS, Local: "Override"}, Attributes: []xml.Attr{{Name: xml.Name{Local: "PartName"}, Value: partURI(a.Name)}, {Name: xml.Name{Local: "ContentType"}, Value: a.ContentType}}})
	}
	if len(nodes) > 0 {
		data, _, e := p.Part(ContentTypesPath)
		if e != nil {
			return nil, e
		}
		doc, e := losslessxml.Parse(data)
		if e != nil {
			return nil, e
		}
		data, e = doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: doc.Elements()[0], Children: nodes}})
		if e != nil {
			return nil, graphEditError("relationship_policy", ContentTypesPath, e.Error())
		}
		plan.patches[ContentTypesPath] = data
	}
	type edgeKey struct{ source, id string }
	edges := map[edgeKey]Edge{}
	for _, e := range graph.Edges {
		edges[edgeKey{e.Source, e.ID}] = e
	}
	seen := map[edgeKey]bool{}
	grouped := map[string][]RelationshipRetarget{}
	for _, edit := range change.Retargets {
		key := edgeKey{edit.Source, edit.ID}
		if seen[key] {
			return nil, graphEditError("ambiguous_target", edit.Source, "duplicate edge edit")
		}
		seen[key] = true
		edge, ok := edges[key]
		if !ok {
			return nil, graphEditError("missing_target", edit.Source, "relationship not found")
		}
		if edge.External {
			return nil, graphEditError("relationship_policy", edit.Source, "external edges cannot be retargeted by this operation")
		}
		parsed, e := url.Parse(edge.Target)
		if e != nil || parsed.Fragment != "" || strings.Contains(edge.Target, "#") {
			return nil, graphEditError("relationship_policy", edit.Source, "fragment-bearing edge requires format policy")
		}
		if err = graphPartName(edit.TargetPart); err != nil {
			return nil, err
		}
		if !p.hasPart(edit.TargetPart) {
			if _, ok := plan.additions[edit.TargetPart]; !ok {
				return nil, graphEditError("missing_target", edit.TargetPart, "new edge target absent")
			}
		}
		if edit.TargetPart == edge.ResolvedPart {
			continue
		} // preserve original relative/escaped spelling on semantic no-op
		name := PackageRelsPath
		if edit.Source != "" {
			name = RelationshipsPathForPart(edit.Source)
		}
		grouped[name] = append(grouped[name], edit)
	}
	for name, edits := range grouped {
		data, _, e := p.Part(name)
		if e != nil {
			return nil, e
		}
		doc, e := losslessxml.Parse(data)
		if e != nil {
			return nil, e
		}
		elements := map[string]losslessxml.Element{}
		for _, node := range doc.Elements()[1:] {
			a, e := attributes(node, "Id", "Type", "Target", "TargetMode")
			if e != nil {
				return nil, graphEditError("relationship_policy", name, e.Error())
			}
			elements[a["Id"]] = node
		}
		attrs := []losslessxml.AttributeEdit{}
		for _, edit := range edits {
			node, ok := elements[edit.ID]
			if !ok {
				return nil, graphEditError("missing_target", name, "relationship element absent")
			}
			attrs = append(attrs, losslessxml.AttributeEdit{Target: node, Name: xml.Name{Local: "Target"}, Value: retargetURI(edit.Source, edit.TargetPart, edges[edgeKey{edit.Source, edit.ID}].Target)})
		}
		data, e = doc.Edit(nil, attrs)
		if e != nil {
			return nil, graphEditError("relationship_policy", name, e.Error())
		}
		plan.patches[name] = data
	}
	if err = p.planGraphAdditions(change, graph, plan); err != nil {
		return nil, err
	}
	if err = p.planGraphRemovals(change, plan); err != nil {
		return nil, err
	}
	// Distinguish a surviving inbound edge from other registry errors. The
	// original graph has already been validated; only explicitly removed or
	// retargeted edges cease to reference the deleted part.
	removed := map[edgeKey]bool{}
	retargeted := map[edgeKey]string{}
	for _, r := range change.Removals {
		removed[edgeKey{r.Source, r.ID}] = true
	}
	for _, r := range change.Retargets {
		retargeted[edgeKey{r.Source, r.ID}] = r.TargetPart
	}
	for _, e := range graph.Edges {
		if !plan.deletions[e.ResolvedPart] || e.External {
			continue
		}
		key := edgeKey{e.Source, e.ID}
		if removed[key] || retargeted[key] != "" && retargeted[key] != e.ResolvedPart {
			continue
		}
		return nil, graphEditError("opc-part-referenced", e.ResolvedPart, "part still has an inbound relationship")
	}
	for _, e := range change.Relationships {
		if plan.deletions[e.TargetPart] {
			return nil, graphEditError("opc-part-referenced", e.TargetPart, "added relationship references deleted part")
		}
	}
	// Validate the complete candidate registry graph before exposing a plan.
	candidate := &Preserved{parts: p.currentParts()}
	for name, data := range plan.additions {
		candidate.parts[name] = data
	}
	for name, data := range plan.patches {
		candidate.parts[name] = data
	}
	for name := range plan.deletions {
		delete(candidate.parts, name)
	}
	if _, err = candidate.Graph(); err != nil {
		return nil, err
	}
	return plan, nil
}

// ApplyGraphPlan commits an already validated plan once. Any intervening actual
// payload change stales the plan, including edits subsequently reverted to source.
// A no-op plan does not consume itself or alter the session generation.
func (p *Preserved) ApplyGraphPlan(plan *GraphPlan) error {
	if plan == nil || plan.owner != p || plan.consumed || plan.generation != p.generation {
		return graphEditError("stale_target", "", "foreign, consumed or stale graph plan")
	}
	if len(plan.additions) == 0 && len(plan.patches) == 0 && len(plan.deletions) == 0 {
		return nil
	}
	if p.added == nil {
		p.added = map[string][]byte{}
	}
	for name, data := range plan.additions {
		p.added[name] = bytes.Clone(data)
	}
	for name, data := range plan.patches {
		if _, ok := p.added[name]; ok {
			p.added[name] = bytes.Clone(data)
			continue
		}
		if bytes.Equal(data, p.parts[name]) {
			delete(p.changed, name)
		} else {
			p.changed[name] = bytes.Clone(data)
		}
	}
	if p.deleted == nil {
		p.deleted = map[string]bool{}
	}
	for name := range plan.deletions {
		delete(p.changed, name)
		delete(p.added, name)
		if _, original := p.parts[name]; original {
			p.deleted[name] = true
		}
	}
	p.generation++
	p.graphEdited = true
	plan.consumed = true
	return nil
}

func (p *Preserved) hasPart(name string) bool {
	_, a := p.parts[name]
	_, b := p.added[name]
	return (a || b) && !p.deleted[name]
}
func (p *Preserved) partNames() []string {
	names := make([]string, 0, len(p.parts)+len(p.added))
	for name := range p.parts {
		if !p.deleted[name] {
			names = append(names, name)
		}
	}
	for name := range p.added {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}
func (p *Preserved) currentParts() map[string][]byte {
	out := make(map[string][]byte, len(p.parts)+len(p.added))
	for name, data := range p.parts {
		out[name] = data
	}
	for name, data := range p.added {
		out[name] = data
	}
	for name, data := range p.changed {
		out[name] = data
	}
	for name := range p.deleted {
		delete(out, name)
	}
	return out
}
