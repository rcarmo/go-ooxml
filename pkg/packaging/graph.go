package packaging

import (
	"encoding/xml"
	"fmt"
	"net/url"
	"path"
	"sort"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// Edge retains the original target spelling and resolved member identity.
// External edges have no ResolvedPart and are never fetched.
type Edge struct {
	Source       string `json:"source"`
	ID           string `json:"id"`
	Type         string `json:"type"`
	Target       string `json:"target"`
	External     bool   `json:"external"`
	ResolvedPart string `json:"resolved_part,omitempty"`
}
type GraphPart struct {
	Name        string `json:"name"`
	ContentType string `json:"content_type"`
	Inbound     int    `json:"inbound"`
}
type Graph struct {
	Parts []GraphPart `json:"parts"`
	Edges []Edge      `json:"edges"`
}

const relNS = "http://schemas.openxmlformats.org/package/2006/relationships"
const ctNS = "http://schemas.openxmlformats.org/package/2006/content-types"

func graphError(part, detail string) error {
	return &Refusal{Kind: "relationship_policy", Operation: "graph", Part: part, Detail: detail}
}

func registry(data []byte, ns, root string) ([]losslessxml.Element, error) {
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	elements := d.Elements()
	if len(elements) == 0 || elements[0].Name() != (xml.Name{Space: ns, Local: root}) {
		return nil, fmt.Errorf("unexpected registry root")
	}
	if len(elements[0].Attributes()) != 0 {
		return nil, fmt.Errorf("unknown registry root attributes")
	}
	if text, _ := elements[0].Text(); strings.TrimSpace(text) != "" {
		return nil, fmt.Errorf("registry mixed text")
	}
	for _, e := range elements[1:] {
		parent, ok := e.Parent()
		if !ok || parent != elements[0] {
			return nil, fmt.Errorf("nested registry extensions unsupported")
		}
		text, leaf := e.Text()
		if !leaf || strings.TrimSpace(text) != "" {
			return nil, fmt.Errorf("registry entry content unsupported")
		}
	}
	return elements[1:], nil
}
func attributes(e losslessxml.Element, allowed ...string) (map[string]string, error) {
	out := map[string]string{}
	for _, a := range e.Attributes() {
		known := false
		for _, name := range allowed {
			if a.Name.Space == "" && a.Name.Local == name {
				known = true
				break
			}
		}
		if !known {
			return nil, fmt.Errorf("unknown attribute %s", a.Name.Local)
		}
		out[a.Name.Local] = a.Value
	}
	return out, nil
}
func sourceForRegistry(name string) (string, error) {
	if name == PackageRelsPath {
		return "", nil
	}
	if path.Base(path.Dir(name)) != "_rels" || !strings.HasSuffix(name, ".rels") || path.Base(name) == ".rels" {
		return "", fmt.Errorf("invalid relationship registry location")
	}
	source := path.Join(path.Dir(path.Dir(name)), strings.TrimSuffix(path.Base(name), ".rels"))
	if err := validateMemberName(source, false); err != nil {
		return "", err
	}
	return source, nil
}
func resolveGraphTarget(source, target string) (string, error) {
	if target == "" || strings.Contains(target, `\`) {
		return "", fmt.Errorf("empty or backslash target")
	}
	u, err := url.Parse(target)
	if err != nil {
		return "", err
	}
	if u.Scheme != "" || u.Host != "" || u.User != nil || u.RawQuery != "" || u.ForceQuery || u.Opaque != "" {
		return "", fmt.Errorf("internal target is not a package URI")
	}
	lower := strings.ToLower(u.EscapedPath())
	if strings.Contains(lower, "%2f") || strings.Contains(lower, "%5c") {
		return "", fmt.Errorf("escaped path separator")
	}
	if u.Path == "" {
		if source == "" {
			return "", fmt.Errorf("empty root target")
		}
		return source, nil
	}
	targetPath := u.Path
	if strings.HasPrefix(targetPath, "/") {
		targetPath = strings.TrimPrefix(targetPath, "/")
	} else {
		targetPath = path.Join(path.Dir(source), targetPath)
	}
	if err := validateMemberName(targetPath, false); err != nil {
		return "", err
	}
	return targetPath, nil
}

// Graph validates the understood OPC registry subset without mutating source.
// Unknown registry extensions refuse; unrelated opaque parts are retained.
// Local edges must resolve. External targets are recorded but never fetched.
func (p *Preserved) Graph() (Graph, error) {
	out := Graph{Parts: []GraphPart{}, Edges: []Edge{}}
	data, _, err := p.Part(ContentTypesPath)
	if err != nil {
		return Graph{}, err
	}
	entries, err := registry(data, ctNS, "Types")
	if err != nil {
		return Graph{}, graphError(ContentTypesPath, err.Error())
	}
	defaults, overrides := map[string]string{}, map[string]string{}
	for _, e := range entries {
		n := e.Name()
		if n.Space != ctNS {
			return Graph{}, graphError(ContentTypesPath, "unknown content-type namespace")
		}
		switch n.Local {
		case "Default":
			a, err := attributes(e, "Extension", "ContentType")
			if err != nil {
				return Graph{}, graphError(ContentTypesPath, err.Error())
			}
			ext := strings.ToLower(a["Extension"])
			if ext == "" || strings.ContainsAny(ext, "./\\") || a["ContentType"] == "" {
				return Graph{}, graphError(ContentTypesPath, "invalid default")
			}
			if _, ok := defaults[ext]; ok {
				return Graph{}, graphError(ContentTypesPath, "duplicate default")
			}
			defaults[ext] = a["ContentType"]
		case "Override":
			a, err := attributes(e, "PartName", "ContentType")
			if err != nil {
				return Graph{}, graphError(ContentTypesPath, err.Error())
			}
			if !strings.HasPrefix(a["PartName"], "/") || a["ContentType"] == "" {
				return Graph{}, graphError(ContentTypesPath, "invalid override")
			}
			name, err := resolveGraphTarget("", a["PartName"])
			if err != nil {
				return Graph{}, graphError(ContentTypesPath, err.Error())
			}
			if _, ok := overrides[name]; ok {
				return Graph{}, graphError(ContentTypesPath, "duplicate override")
			}
			if _, ok := p.parts[name]; !ok {
				return Graph{}, graphError(ContentTypesPath, "override targets absent part")
			}
			overrides[name] = a["ContentType"]
		default:
			return Graph{}, graphError(ContentTypesPath, "unknown content-type entry")
		}
	}
	names := make([]string, 0, len(p.parts))
	for name := range p.parts {
		names = append(names, name)
	}
	sort.Strings(names)
	inbound := map[string]int{}
	for _, name := range names {
		if name == ContentTypesPath {
			continue
		}
		ct := overrides[name]
		if ct == "" {
			ct = defaults[strings.ToLower(strings.TrimPrefix(path.Ext(name), "."))]
		}
		if ct == "" {
			return Graph{}, graphError(name, "missing content type")
		}
		out.Parts = append(out.Parts, GraphPart{Name: name, ContentType: ct})
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		source, err := sourceForRegistry(name)
		if err != nil {
			return Graph{}, graphError(name, err.Error())
		}
		if source != "" {
			if _, ok := p.parts[source]; !ok {
				return Graph{}, graphError(name, "relationship source missing")
			}
		}
		data, _, err := p.Part(name)
		if err != nil {
			return Graph{}, err
		}
		entries, err := registry(data, relNS, "Relationships")
		if err != nil {
			return Graph{}, graphError(name, err.Error())
		}
		seen := map[string]bool{}
		for _, e := range entries {
			if e.Name() != (xml.Name{Space: relNS, Local: "Relationship"}) {
				return Graph{}, graphError(name, "unknown relationship entry")
			}
			a, err := attributes(e, "Id", "Type", "Target", "TargetMode")
			if err != nil {
				return Graph{}, graphError(name, err.Error())
			}
			if a["Id"] == "" || a["Type"] == "" || a["Target"] == "" || seen[a["Id"]] {
				return Graph{}, graphError(name, "missing fields or duplicate relationship ID")
			}
			seen[a["Id"]] = true
			edge := Edge{Source: source, ID: a["Id"], Type: a["Type"], Target: a["Target"]}
			switch a["TargetMode"] {
			case "External":
				edge.External = true
			case "", "Internal":
				edge.ResolvedPart, err = resolveGraphTarget(source, a["Target"])
				if err != nil {
					return Graph{}, graphError(name, err.Error())
				}
				if _, ok := p.parts[edge.ResolvedPart]; !ok {
					return Graph{}, graphError(name, "target missing: "+edge.ResolvedPart)
				}
				inbound[edge.ResolvedPart]++
			default:
				return Graph{}, graphError(name, "unknown TargetMode")
			}
			out.Edges = append(out.Edges, edge)
		}
	}
	for i := range out.Parts {
		out.Parts[i].Inbound = inbound[out.Parts[i].Name]
	}
	sort.Slice(out.Edges, func(i, j int) bool {
		a, b := out.Edges[i], out.Edges[j]
		if a.Source == b.Source {
			return a.ID < b.ID
		}
		return a.Source < b.Source
	})
	return out, nil
}
