package acceptance

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// contract20CreationPolicy is an independent member/role/edge allowlist for
// the sealed creation profile. Names are discovered through content types and
// resolved relationships, not fixed to a producer's slide numbering.
func contract20CreationPolicy(members map[string][]byte, g packaging.Graph, slides int) error {
	types := map[string]string{
		packaging.ContentTypePresentation: "presentation", packaging.ContentTypeSlide: "slide", packaging.ContentTypeSlideMaster: "slideMaster", packaging.ContentTypeSlideLayout: "slideLayout", packaging.ContentTypeTheme: "theme",
		packaging.ContentTypePresentationProps: "presentationProperties", packaging.ContentTypePresentationViewProps: "viewProperties",
		"application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml": "tableStyleDefinitions",
		packaging.ContentTypeCoreProps: "coreProperties", packaging.ContentTypeExtendedProps: "applicationProperties",
	}
	required := map[string]int{"presentation": 1, "slide": slides, "slideMaster": 1, "slideLayout": 1, "theme": 1}
	roles := map[string]string{}
	parts := map[string]packaging.GraphPart{}
	counts := map[string]int{}
	registries := map[string]bool{}
	for _, p := range g.Parts {
		if _, dup := parts[p.Name]; dup {
			return fmt.Errorf("duplicate created part %s", p.Name)
		}
		parts[p.Name] = p
		if _, ok := members[p.Name]; !ok {
			return fmt.Errorf("graph part absent %s", p.Name)
		}
		if strings.HasSuffix(p.Name, ".rels") {
			if p.ContentType != packaging.ContentTypeRelationships {
				return fmt.Errorf("wrong relationship type %s", p.Name)
			}
			registries[p.Name] = true
			continue
		}
		role, ok := types[p.ContentType]
		if !ok {
			return fmt.Errorf("unknown created support role %s %s", p.Name, p.ContentType)
		}
		roles[p.Name] = role
		counts[role]++
	}
	if len(parts)+1 != len(members) {
		return fmt.Errorf("member/graph inventory differs: %d %d", len(parts)+1, len(members))
	}
	if _, ok := members[packaging.ContentTypesPath]; !ok {
		return fmt.Errorf("content type registry absent")
	}
	ct, err := losslessxml.Parse(members[packaging.ContentTypesPath])
	if err != nil || len(ct.Elements()) == 0 || ct.Elements()[0].Name() != (xml.Name{Space: packaging.NSContentTypes, Local: "Types"}) {
		return fmt.Errorf("invalid content type registry: %v", err)
	}
	overrides := map[string]string{}
	defaults := map[string]string{}
	for _, node := range ct.Elements()[1:] {
		attributes := map[string]string{}
		for _, a := range node.Attributes() {
			if a.Name.Space != "" {
				return fmt.Errorf("foreign content type attribute")
			}
			attributes[a.Name.Local] = a.Value
		}
		switch node.Name() {
		case xml.Name{Space: packaging.NSContentTypes, Local: "Override"}:
			part := strings.TrimPrefix(attributes["PartName"], "/")
			if attributes["PartName"] != "/"+part || overrides[part] != "" || parts[part].Name == "" || attributes["ContentType"] != parts[part].ContentType {
				return fmt.Errorf("invalid content type override %s", part)
			}
			overrides[part] = attributes["ContentType"]
		case xml.Name{Space: packaging.NSContentTypes, Local: "Default"}:
			ext := attributes["Extension"]
			if ext == "" || defaults[ext] != "" || attributes["ContentType"] == "" {
				return fmt.Errorf("invalid content type default %q", ext)
			}
			defaults[ext] = attributes["ContentType"]
		default:
			return fmt.Errorf("unknown content type entry %s", node.Name().Local)
		}
	}
	for part, p := range parts {
		if p.ContentType != overrides[part] {
			ext := part[strings.LastIndex(part, ".")+1:]
			if defaults[ext] != p.ContentType {
				return fmt.Errorf("content type not justified by registry: %s", part)
			}
		}
	}
	for role, n := range required {
		if counts[role] != n {
			return fmt.Errorf("created role %s count %d != %d", role, counts[role], n)
		}
	}
	for role, n := range counts {
		if _, ok := required[role]; !ok && n > 1 {
			return fmt.Errorf("optional role %s count %d", role, n)
		}
	}
	type rule struct {
		source, target, kind string
		required             bool
	}
	rules := []rule{
		{"packageRoot", "presentation", packaging.RelTypeOfficeDocument, true},
		{"presentation", "slide", packaging.RelTypeSlide, true},
		{"presentation", "slideMaster", packaging.RelTypeSlideMaster, true},
		{"slide", "slideLayout", packaging.RelTypeSlideLayout, true},
		{"slideLayout", "slideMaster", packaging.RelTypeSlideMaster, true},
		{"slideMaster", "slideLayout", packaging.RelTypeSlideLayout, true},
		{"slideMaster", "theme", packaging.RelTypeTheme, true},
		{"presentation", "theme", packaging.RelTypeTheme, false},
		{"presentation", "presentationProperties", packaging.RelTypePresProps, false},
		{"presentation", "viewProperties", packaging.RelTypeViewProps, false},
		{"presentation", "tableStyleDefinitions", packaging.RelTypeTableStyles, false},
		{"packageRoot", "coreProperties", packaging.RelTypeCoreProps, false},
		{"packageRoot", "applicationProperties", packaging.RelTypeExtendedProps, false},
	}
	key := func(source, target, kind string) string { return source + ">" + target + ">" + kind }
	allowed := map[string]rule{}
	for _, r := range rules {
		allowed[key(r.source, r.target, r.kind)] = r
	}
	seen := map[string]int{}
	actualRegistries := map[string]bool{}
	for _, e := range g.Edges {
		if e.External || e.ResolvedPart == "" {
			return fmt.Errorf("external or unresolved edge %+v", e)
		}
		owner := "packageRoot"
		if e.Source != "" {
			owner = roles[e.Source]
		}
		target := roles[e.ResolvedPart]
		k := key(owner, target, e.Type)
		r, ok := allowed[k]
		if !ok {
			return fmt.Errorf("unknown created edge %+v roles %s>%s", e, owner, target)
		}
		seen[k]++
		limit := 1
		if r.target == "slide" {
			limit = slides
		}
		if r.source == "slide" {
			limit = slides
		}
		if seen[k] > limit {
			return fmt.Errorf("duplicate created edge %s", k)
		}
		reg := packaging.PackageRelsPath
		if e.Source != "" {
			reg = packaging.RelationshipsPathForPart(e.Source)
		}
		actualRegistries[reg] = true
	}
	for source, role := range roles {
		if role != "slide" {
			continue
		}
		links := 0
		for _, edge := range g.Edges {
			if edge.Source == source && edge.Type == packaging.RelTypeSlideLayout && edge.ResolvedPart != "" && roles[edge.ResolvedPart] == "slideLayout" {
				links++
			}
		}
		if links != 1 {
			return fmt.Errorf("slide %s has %d layout edges", source, links)
		}
	}
	for _, r := range rules {
		k := key(r.source, r.target, r.kind)
		want := 0
		if r.required {
			want = 1
			if r.target == "slide" || r.source == "slide" {
				want = slides
			}
		}
		if r.required && seen[k] != want {
			return fmt.Errorf("required created edge %s: %d != %d", k, seen[k], want)
		}
		if !r.required && r.target != "theme" && seen[k] != counts[r.target] {
			return fmt.Errorf("orphan optional role %s", r.target)
		}
		if !r.required && seen[k] > 1 {
			return fmt.Errorf("duplicate optional edge %s", k)
		}
	}
	if !reflect.DeepEqual(registries, actualRegistries) {
		return fmt.Errorf("orphan/missing relationship member %v != %v", registries, actualRegistries)
	}
	inbound := map[string]int{}
	for _, e := range g.Edges {
		inbound[e.ResolvedPart]++
	}
	for part, role := range roles {
		if inbound[part] == 0 || parts[part].Inbound != inbound[part] {
			return fmt.Errorf("unreachable support part %s (%s)", part, role)
		}
	}
	// Independently cross-check slide/master list identities and title-layout
	// placeholder inventory against parsed graph edges, not package filenames.
	var main, master, layout string
	for n, r := range roles {
		switch r {
		case "presentation":
			main = n
		case "slideMaster":
			master = n
		case "slideLayout":
			layout = n
		}
	}
	presentationDoc, err := losslessxml.Parse(members[main])
	if err != nil {
		return err
	}
	masterDoc, err := losslessxml.Parse(members[master])
	if err != nil {
		return err
	}
	layoutDoc, err := losslessxml.Parse(members[layout])
	if err != nil {
		return err
	}
	for doc, expected := range map[*losslessxml.Document]string{presentationDoc: "presentation", masterDoc: "sldMaster", layoutDoc: "sldLayout"} {
		if len(doc.Elements()) == 0 || doc.Elements()[0].Name() != (xml.Name{Space: packaging.NSPresentationML, Local: expected}) {
			return fmt.Errorf("wrong created root %s", expected)
		}
	}
	pIDs, mIDs := map[string]bool{}, map[string]bool{}
	slideInts := map[string]bool{}
	layoutInts := map[string]bool{}
	placeholders := map[string]int{}
	for _, n := range presentationDoc.Elements() {
		if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldId"}) && n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldMasterId"}) {
			continue
		}
		id, rid := "", ""
		for _, a := range n.Attributes() {
			if a.Name.Space == "" && a.Name.Local == "id" {
				id = a.Value
			}
			if a.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}) {
				rid = a.Value
			}
		}
		value, e := strconv.ParseUint(id, 10, 64)
		if e != nil || value == 0 || rid == "" {
			return fmt.Errorf("invalid presentation inventory ID %q/%q", id, rid)
		}
		if n.Name().Local == "sldId" {
			if pIDs[rid] || slideInts[id] {
				return fmt.Errorf("duplicate slide inventory")
			}
			pIDs[rid] = true
			slideInts[id] = true
		} else {
			if mIDs[rid] {
				return fmt.Errorf("duplicate master inventory")
			}
			mIDs[rid] = true
		}
	}
	if len(pIDs) != slides || len(mIDs) != 1 {
		return fmt.Errorf("slide/master list size %d/%d", len(pIDs), len(mIDs))
	}
	for _, e := range g.Edges {
		if e.Source != main {
			continue
		}
		if e.Type == packaging.RelTypeSlide && !pIDs[e.ID] || e.Type == packaging.RelTypeSlideMaster && !mIDs[e.ID] {
			return fmt.Errorf("unlisted graph target %+v", e)
		}
	}
	for _, n := range masterDoc.Elements() {
		if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldLayoutId"}) {
			continue
		}
		id, rid := "", ""
		for _, a := range n.Attributes() {
			if a.Name.Space == "" && a.Name.Local == "id" {
				id = a.Value
			}
			if a.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}) {
				rid = a.Value
			}
		}
		value, e := strconv.ParseUint(id, 10, 64)
		if e != nil || value == 0 || rid == "" || layoutInts[id] {
			return fmt.Errorf("invalid layout inventory ID %q/%q", id, rid)
		}
		layoutInts[id] = true
		found := false
		for _, edge := range g.Edges {
			if edge.Source == master && edge.Type == packaging.RelTypeSlideLayout && edge.ID == rid && edge.ResolvedPart == layout {
				found = true
			}
		}
		if !found {
			return fmt.Errorf("unresolved master layout ID %q", rid)
		}
	}
	if len(layoutInts) != 1 {
		return fmt.Errorf("layout list count %d", len(layoutInts))
	}
	root := layoutDoc.Elements()[0]
	layoutType := ""
	for _, a := range root.Attributes() {
		if a.Name.Space == "" && a.Name.Local == "type" {
			layoutType = a.Value
		}
	}
	if layoutType != "title" {
		return fmt.Errorf("wrong title layout %q", layoutType)
	}
	for _, n := range layoutDoc.Elements() {
		if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "ph"}) {
			continue
		}
		for _, a := range n.Attributes() {
			if a.Name.Space == "" && a.Name.Local == "type" {
				placeholders[a.Value]++
			}
		}
	}
	if placeholders["ctrTitle"] != 1 || placeholders["subTitle"] != 1 {
		return fmt.Errorf("title/subtitle layout placeholders %v", placeholders)
	}
	for slide, role := range roles {
		if role != "slide" {
			continue
		}
		d, err := losslessxml.Parse(members[slide])
		if err != nil {
			return err
		}
		if len(d.Elements()) == 0 || d.Elements()[0].Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sld"}) {
			return fmt.Errorf("created slide root invalid %s", slide)
		}
		ownedIDs := map[string]bool{}
		found := map[string]int{}
		for _, n := range d.Elements() {
			if n.Name() == (xml.Name{Space: packaging.NSPresentationML, Local: "cNvPr"}) {
				for _, a := range n.Attributes() {
					if a.Name.Space == "" && a.Name.Local == "id" {
						if ownedIDs[a.Value] {
							return fmt.Errorf("duplicate slide object ID %s", a.Value)
						}
						ownedIDs[a.Value] = true
					}
				}
			}
			if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "ph"}) {
				continue
			}
			for _, a := range n.Attributes() {
				if a.Name.Space == "" && a.Name.Local == "type" {
					found[a.Value]++
				}
			}
		}
		if found["ctrTitle"] != 1 || found["subTitle"] != 1 {
			return fmt.Errorf("authored slide placeholder count %v", found)
		}
		if bytes.Contains(members[slide], []byte(`<p:notes`)) {
			return fmt.Errorf("unrequested notes on slide")
		}
	}
	return nil
}
