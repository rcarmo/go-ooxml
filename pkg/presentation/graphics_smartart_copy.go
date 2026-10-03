package presentation

import (
	"crypto/rand"
	"encoding/xml"
	"errors"
	"fmt"
	"math"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type SmartArtCopyReceipt struct {
	ShapeID      uint32            `json:"shapeId"`
	PartName     string            `json:"partName"`
	PartMap      map[string]string `json:"partMap"`
	ModelIDMap   map[string]string `json:"modelIdMap"`
	DrawingIDMap map[string]uint32 `json:"drawingIdMap"`
}

var graphicsGUID = regexp.MustCompile(`^\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}$`)
var graphicsDiagramMIME = regexp.MustCompile(`^application/vnd\.openxmlformats-officedocument\.drawingml\.diagram(?:Data|Layout|Style|Colors)\+xml$`)

func graphicsFreshGUID() (string, error) {
	var b [16]byte
	if _, e := rand.Read(b[:]); e != nil {
		return "", e
	}
	b[6] = (b[6] & 15) | 64
	b[8] = (b[8] & 63) | 128
	return fmt.Sprintf("{%X-%X-%X-%X-%X}", b[0:4], b[4:6], b[6:8], b[8:10], b[10:16]), nil
}
func graphicsRelativePart(owner, target string) string {
	from := strings.Split(path.Dir(owner), "/")
	to := strings.Split(target, "/")
	same := 0
	for same < len(from) && same < len(to) && from[same] == to[same] {
		same++
	}
	parts := []string{}
	for i := same; i < len(from); i++ {
		if from[i] != "." {
			parts = append(parts, "..")
		}
	}
	parts = append(parts, to[same:]...)
	return strings.Join(parts, "/")
}
func graphicsBoundFrame(frame losslessxml.Element) ([]byte, error) {
	raw := frame.Raw()
	end := bytesOpeningEnd(raw)
	if end < 1 {
		return nil, graphicsSmartArtUnsupported("frame opening tag")
	}
	existing := map[string]bool{}
	opening := string(raw[:end])
	matches := regexp.MustCompile(`\s+xmlns(?::([A-Za-z_][\w.-]*))?\s*=`).FindAllStringSubmatch(opening, -1)
	for _, m := range matches {
		existing[m[1]] = true
	}
	prefixes := []string{}
	bindings := frame.Namespaces()
	for p := range bindings {
		if p != "xml" && !existing[p] {
			prefixes = append(prefixes, p)
		}
	}
	sort.Strings(prefixes)
	addition := ""
	for _, p := range prefixes {
		v, e := graphicsXMLString(bindings[p])
		if e != nil {
			return nil, e
		}
		k := "xmlns"
		if p != "" {
			k += ":" + p
		}
		addition += " " + k + `="` + v + `"`
	}
	at := end - 1
	if at > 0 && raw[at-1] == '/' {
		at--
	}
	result := append(append(append([]byte{}, raw[:at]...), []byte(addition)...), raw[at:]...)
	return result, nil
}
func bytesOpeningEnd(raw []byte) int { return fmtOpeningEnd(raw) }

// CopySmartArtFrom clones one admitted frame/closure in a single validated graph plan.
func (s *EditSession) CopySmartArtFrom(targetPart string, source *EditSession, sourcePart string, id uint32) (SmartArtCopyReceipt, error) {
	var none SmartArtCopyReceipt
	if source == nil || !s.slides[targetPart] || !source.slides[sourcePart] || id == 0 || id > math.MaxInt32 {
		return none, graphicsSmartArtUnsupported("copy ownership/identity")
	}
	if e := s.manipulationProtection(); e != nil {
		return none, editRefusal("PPTX_PROTECTED", "destination protection")
	}
	infos, e := source.InspectSmartArt(sourcePart)
	if e != nil {
		return none, e
	}
	var info *SmartArtInfo
	for i := range infos {
		if infos[i].ShapeID == id {
			info = &infos[i]
		}
	}
	if info == nil {
		return none, graphicsSmartArtUnsupported("source frame absent")
	}
	slide, _, e := source.pkg.Part(sourcePart)
	if e != nil {
		return none, e
	}
	slideDoc, e := losslessxml.Parse(slide)
	if e != nil {
		return none, e
	}
	var frame losslessxml.Element
	for _, n := range slideDoc.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "graphicFrame") {
			continue
		}
		nv, er := graphicsSmartOne(slideDoc, n, packaging.NSPresentationML, "nvGraphicFramePr")
		if er != nil {
			return none, er
		}
		props, er := graphicsSmartOne(slideDoc, nv, packaging.NSPresentationML, "cNvPr")
		if er != nil {
			return none, er
		}
		raw, _ := graphicsAttr(props, "", "id")
		v, _ := graphicsNumber(raw, 1, math.MaxInt32)
		if uint32(v) == id {
			frame = n
			break
		}
	}
	parent, ok := frame.Parent()
	if !ok || parent.Name() != name(packaging.NSPresentationML, "spTree") {
		return none, graphicsSmartArtUnsupported("direct source frame required")
	}
	mapAppendError := func(e error) error {
		var r *packaging.Refusal
		if errors.As(e, &r) && r.Kind == "PPTX_PICTURE_UNSUPPORTED" {
			return graphicsSmartArtUnsupported(r.Detail)
		}
		return e
	}
	if _, _, e = graphicsShapeAppend(slide); e != nil {
		return none, mapAppendError(e)
	}
	target, hash, e := s.pkg.Part(targetPart)
	if e != nil {
		return none, e
	}
	nextID, at, e := graphicsShapeAppend(target)
	if e != nil {
		return none, mapAppendError(e)
	}
	partTypes := map[string]string{}
	for _, p := range info.Parts {
		partTypes[p.PartName] = p.ContentType
	}
	expectedTypes := map[string]string{graphicsRelationNS + "/diagramData": "diagramData", graphicsRelationNS + "/diagramLayout": "diagramLayout", graphicsRelationNS + "/diagramQuickStyle": "diagramStyle", graphicsRelationNS + "/diagramColors": "diagramColors"}
	for _, edge := range info.Edges {
		allowed := edge.Type == graphicsDiagramDrawingRel
		for _, role := range []string{"diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors", "image", "hyperlink"} {
			if edge.Type == graphicsRelationNS+"/"+role {
				allowed = true
			}
		}
		if strings.Contains(edge.Target, "#") || !allowed || edge.External && edge.Type != packaging.RelTypeImage && edge.Type != graphicsRelationNS+"/hyperlink" {
			return none, graphicsSmartArtUnsupported("dependency role/target")
		}
		if !edge.External {
			if edge.PartName == nil {
				return none, graphicsSmartArtUnsupported("dependency absent")
			}
			mime := partTypes[*edge.PartName]
			if edge.Type == packaging.RelTypeImage && !strings.HasPrefix(mime, "image/") || edge.Type == graphicsRelationNS+"/hyperlink" {
				return none, graphicsSmartArtUnsupported("dependency role/type")
			}
			if role := expectedTypes[edge.Type]; role != "" && mime != "application/vnd.openxmlformats-officedocument.drawingml."+role+"+xml" {
				return none, graphicsSmartArtUnsupported("dependency role MIME")
			}
		}
	}
	graph, e := s.pkg.Graph()
	if e != nil {
		return none, e
	}
	used := map[string]bool{}
	for _, p := range graph.Parts {
		used[strings.ToLower(p.Name)] = true
	}
	receipt := SmartArtCopyReceipt{nextID, targetPart, map[string]string{}, map[string]string{}, map[string]uint32{}}
	allocate := func(template string) (string, error) {
		for i := 1; i <= len(used)+1; i++ {
			p := fmt.Sprintf(template, i)
			rp := packaging.RelationshipsPathForPart(p)
			if !used[strings.ToLower(p)] && !used[strings.ToLower(rp)] {
				used[strings.ToLower(p)] = true
				used[strings.ToLower(rp)] = true
				return p, nil
			}
		}
		return "", graphicsSmartArtUnsupported("part name exhaustion")
	}
	sourceGraph, e := source.pkg.Graph()
	if e != nil {
		return none, e
	}
	sourceEdges := map[string][]packaging.Edge{}
	for _, edge := range sourceGraph.Edges {
		sourceEdges[edge.Source] = append(sourceEdges[edge.Source], edge)
	}
	values := map[string][]byte{}
	docs := map[string]*losslessxml.Document{}
	for _, p := range info.Parts {
		template := "ppt/diagrams/smartArt%d.xml"
		if strings.HasPrefix(p.ContentType, "image/") {
			if len(sourceEdges[p.PartName]) != 0 {
				return none, graphicsSmartArtUnsupported("image-owned dependencies")
			}
			ext := path.Ext(p.PartName)
			if !regexp.MustCompile(`^\.[A-Za-z0-9]+$`).MatchString(ext) {
				return none, graphicsSmartArtUnsupported("image extension")
			}
			template = "ppt/media/smartArt%d" + ext
		} else if !graphicsDiagramMIME.MatchString(p.ContentType) && p.ContentType != "application/vnd.ms-office.drawingml.diagramDrawing+xml" {
			return none, graphicsSmartArtUnsupported("closure content type")
		}
		mapped, er := allocate(template)
		if er != nil {
			return none, er
		}
		receipt.PartMap[p.PartName] = mapped
		raw, _, er := source.pkg.Part(p.PartName)
		if er != nil {
			return none, er
		}
		values[p.PartName] = raw
		if strings.HasPrefix(p.ContentType, "image/") {
			continue
		}
		d, er := losslessxml.Parse(raw)
		if er != nil {
			return none, er
		}
		docs[p.PartName] = d
		if p.ContentType == "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" {
			for _, n := range d.Elements() {
				if n.Name() != name(graphicsDiagramNS, "pt") && n.Name() != name(graphicsDiagramNS, "cxn") {
					continue
				}
				old, _ := graphicsAttr(n, "", "modelId")
				if !graphicsGUID.MatchString(old) {
					return none, graphicsSmartArtUnsupported("instance model GUID")
				}
				if _, exists := receipt.ModelIDMap[old]; exists {
					return none, graphicsSmartArtUnsupported("duplicate model definition")
				}
				receipt.ModelIDMap[old] = ""
			}
		}
	}
	reserved := map[string]bool{}
	for old := range receipt.ModelIDMap {
		reserved[strings.ToUpper(old)] = true
	}
	for old := range receipt.ModelIDMap {
		for {
			fresh, er := graphicsFreshGUID()
			if er != nil {
				return none, er
			}
			if !reserved[fresh] {
				receipt.ModelIDMap[old] = fresh
				reserved[fresh] = true
				break
			}
		}
	}
	targetRIDs := map[string]bool{}
	for _, edge := range graph.Edges {
		if edge.Source == targetPart {
			targetRIDs[edge.ID] = true
		}
	}
	nextRID := func() string {
		for i := 1; ; i++ {
			rid := "rId" + strconv.Itoa(i)
			if !targetRIDs[rid] {
				targetRIDs[rid] = true
				return rid
			}
		}
	}
	change := packaging.GraphMutation{}
	drawingRIDs := map[string]string{}
	for _, edge := range info.Edges {
		if edge.Owner == sourcePart && edge.Type == graphicsDiagramDrawingRel {
			rid := nextRID()
			drawingRIDs[edge.RelationshipID] = rid
			change.Relationships = append(change.Relationships, packaging.RelationshipAddition{Source: targetPart, ID: rid, Type: edge.Type, TargetPart: receipt.PartMap[*edge.PartName]})
		}
	}
	known := map[string]bool{}
	for _, k := range []string{"modelId", "srcId", "destId", "cxnId", "parTransId", "sibTransId", "presId", "presAssocID"} {
		known[k] = true
	}
	for part, d := range docs {
		es := d.Elements()
		instance := es[0].Name().Space == graphicsPersistDiagramNS || es[0].Name() == name(graphicsDiagramNS, "dataModel")
		nodes := []losslessxml.Element{}
		for _, n := range es {
			if n.Name() == name(graphicsPersistDiagramNS, "cNvPr") {
				nodes = append(nodes, n)
			}
		}
		localIDs := map[string]uint32{}
		nodeIDs := map[losslessxml.Element]uint32{}
		maximum := int64(0)
		zeroCount := 0
		for _, n := range nodes {
			raw, _ := graphicsAttr(n, "", "id")
			v, er := graphicsNumber(raw, 0, math.MaxInt32)
			if er != nil || strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") {
				return none, graphicsSmartArtUnsupported("drawing ID")
			}
			if raw != "0" {
				if _, ok := localIDs[raw]; ok {
					return none, graphicsSmartArtUnsupported("duplicate drawing ID")
				}
			} else {
				zeroCount++
			}
			localIDs[raw] = 0
			if v > maximum {
				maximum = v
			}
		}
		if maximum+int64(len(nodes)) > math.MaxInt32 {
			return none, graphicsSmartArtUnsupported("drawing identity overflow")
		}
		zeroIndex := 0
		for _, n := range nodes {
			old, _ := graphicsAttr(n, "", "id")
			maximum++
			v := uint32(maximum)
			localIDs[old] = v
			nodeIDs[n] = v
			key := part + "#" + old
			if old == "0" && zeroCount > 1 {
				key += "@" + strconv.Itoa(zeroIndex)
				zeroIndex++
			}
			receipt.DrawingIDMap[key] = v
		}
		edits := []losslessxml.AttributeEdit{}
		for _, n := range es {
			changes := map[xml.Name]string{}
			for _, a := range n.Attributes() {
				if instance && a.Name.Space == "" && known[a.Name.Local] {
					if a.Value == "" {
						continue
					}
					if a.Name.Local == "presId" && strings.HasPrefix(a.Value, "urn:microsoft.com/office/officeart/2005/8/layout/") {
						continue
					}
					mapped, ok := receipt.ModelIDMap[a.Value]
					if !ok {
						return none, graphicsSmartArtUnsupported("dangling model reference")
					}
					changes[a.Name] = mapped
				} else if instance && graphicsGUID.MatchString(a.Value) && a.Name.Local != "uniqueId" && a.Name.Local != "uri" {
					return none, graphicsSmartArtUnsupported("unknown GUID reference")
				}
			}
			if n.Name() == name(graphicsPersistDiagramNS, "cNvPr") {
				changes[name("", "id")] = strconv.FormatUint(uint64(nodeIDs[n]), 10)
			}
			if n.Name() == name(packaging.NSDrawingML, "stCxn") || n.Name() == name(packaging.NSDrawingML, "endCxn") {
				old, _ := graphicsAttr(n, "", "id")
				mapped, ok := localIDs[old]
				if !ok || old == "0" && zeroCount > 1 {
					return none, graphicsSmartArtUnsupported("ambiguous/dangling drawing attachment")
				}
				changes[name("", "id")] = strconv.FormatUint(uint64(mapped), 10)
			}
			if n.Name() == name(graphicsPersistDiagramNS, "dataModelExt") {
				old, _ := graphicsAttr(n, "", "relId")
				if rid := drawingRIDs[old]; rid != "" {
					changes[name("", "relId")] = rid
				}
			}
			for k, v := range changes {
				edits = append(edits, losslessxml.AttributeEdit{Target: n, Name: k, Value: v})
			}
		}
		raw, er := d.Edit(nil, edits)
		if er != nil {
			return none, er
		}
		values[part] = raw
	}
	for _, p := range info.Parts {
		change.Additions = append(change.Additions, packaging.PartAddition{Name: receipt.PartMap[p.PartName], ContentType: p.ContentType, Data: values[p.PartName]})
		raw, _, er := source.pkg.Part(packaging.RelationshipsPathForPart(p.PartName))
		if er != nil {
			if len(sourceEdges[p.PartName]) != 0 {
				return none, er
			}
			continue
		}
		d, er := losslessxml.Parse(raw)
		if er != nil {
			return none, er
		}
		edits := []losslessxml.AttributeEdit{}
		for _, n := range d.Elements()[1:] {
			id, _ := graphicsAttr(n, "", "Id")
			var edge *SmartArtEdge
			for i := range info.Edges {
				if info.Edges[i].Owner == p.PartName && info.Edges[i].RelationshipID == id {
					edge = &info.Edges[i]
					break
				}
			}
			if edge == nil {
				return none, graphicsSmartArtUnsupported("uninspected registry edge")
			}
			if !edge.External {
				edits = append(edits, losslessxml.AttributeEdit{Target: n, Name: name("", "Target"), Value: graphicsRelativePart(receipt.PartMap[p.PartName], receipt.PartMap[*edge.PartName])})
			}
		}
		retargeted, er := d.Edit(nil, edits)
		if er != nil {
			return none, er
		}
		change.Registries = append(change.Registries, packaging.RelationshipRegistryAddition{Owner: receipt.PartMap[p.PartName], Data: retargeted})
	}
	roleMap := map[string]string{}
	for _, r := range info.Roots {
		rid := nextRID()
		roleMap[r.RelationshipID] = rid
		var typ string
		for _, edge := range sourceEdges[sourcePart] {
			if edge.ID == r.RelationshipID {
				typ = edge.Type
			}
		}
		change.Relationships = append(change.Relationships, packaging.RelationshipAddition{Source: targetPart, ID: rid, Type: typ, TargetPart: receipt.PartMap[r.PartName]})
	}
	bound, e := graphicsBoundFrame(frame)
	if e != nil {
		return none, e
	}
	fd, e := losslessxml.Parse(bound)
	if e != nil {
		return none, e
	}
	var identity, leaf losslessxml.Element
	for _, n := range fd.Elements() {
		if n.Name() == name(packaging.NSPresentationML, "cNvPr") {
			identity = n
		}
		if n.Name() == name(graphicsDiagramNS, "relIds") {
			leaf = n
		}
	}
	edits := []losslessxml.AttributeEdit{{Target: identity, Name: name("", "id"), Value: strconv.FormatUint(uint64(nextID), 10)}}
	for _, a := range leaf.Attributes() {
		if a.Name.Space != packaging.NSDocumentRelationships {
			continue
		}
		rid, ok := roleMap[a.Value]
		if !ok {
			return none, graphicsSmartArtUnsupported("frame role remap")
		}
		edits = append(edits, losslessxml.AttributeEdit{Target: leaf, Name: a.Name, Value: rid})
	}
	copied, e := fd.Edit(nil, edits)
	if e != nil {
		return none, e
	}
	next := append(append(append([]byte{}, target[:at]...), copied...), target[at:]...)
	change.Replacements = append(change.Replacements, packaging.Replacement{Part: targetPart, ExpectedSHA256: hash, Data: next})
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
