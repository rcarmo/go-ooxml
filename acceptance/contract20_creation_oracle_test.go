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

// creationSlideOrder derives each slide part from the saved presentation list
// and resolved OPC edges, never from a producer's slide numbering convention.
func creationSlideOrder(members map[string][]byte, graph packaging.Graph) ([]string, error) {
	data, ok := members[packaging.PresentationPath]
	if !ok {
		return nil, fmt.Errorf("presentation member missing")
	}
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	p := xml.Name{Space: packaging.NSPresentationML, Local: "presentation"}
	listName := xml.Name{Space: packaging.NSPresentationML, Local: "sldIdLst"}
	slideName := xml.Name{Space: packaging.NSPresentationML, Local: "sldId"}
	if len(doc.Elements()) == 0 || doc.Elements()[0].Name() != p {
		return nil, fmt.Errorf("presentation root invalid")
	}
	listCount := 0
	var list losslessxml.Element
	for _, node := range doc.Elements() {
		if node.Name() == listName {
			listCount++
			list = node
		}
	}
	if listCount != 1 {
		return nil, fmt.Errorf("slide-list count %d", listCount)
	}
	edges := map[string]packaging.Edge{}
	for _, edge := range graph.Edges {
		if edge.Source == packaging.PresentationPath && edge.Type == packaging.RelTypeSlide {
			if _, ok := edges[edge.ID]; ok {
				return nil, fmt.Errorf("duplicate slide relationship %q", edge.ID)
			}
			edges[edge.ID] = edge
		}
	}
	types := map[string]string{}
	for _, part := range graph.Parts {
		types[part.Name] = part.ContentType
	}
	out := []string{}
	seenRID := map[string]bool{}
	seenID := map[uint64]bool{}
	seenPart := map[string]bool{}
	for _, node := range doc.Elements() {
		if node.Name() != slideName {
			continue
		}
		parent, ok := node.Parent()
		if !ok || parent != list {
			return nil, fmt.Errorf("slide inventory owner invalid")
		}
		id, rid := "", ""
		for _, attr := range node.Attributes() {
			if attr.Name == (xml.Name{Local: "id"}) {
				id = attr.Value
			}
			if attr.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}) {
				rid = attr.Value
			}
		}
		n, e := strconv.ParseUint(id, 10, 32)
		if e != nil || n < 256 || rid == "" || seenID[n] || seenRID[rid] {
			return nil, fmt.Errorf("duplicate/invalid slide identity")
		}
		seenID[n], seenRID[rid] = true, true
		edge, ok := edges[rid]
		if !ok || edge.External || edge.ResolvedPart == "" || types[edge.ResolvedPart] != packaging.ContentTypeSlide || seenPart[edge.ResolvedPart] {
			return nil, fmt.Errorf("unresolved/duplicate slide part %q", rid)
		}
		if _, ok := members[edge.ResolvedPart]; !ok {
			return nil, fmt.Errorf("listed slide absent %s", edge.ResolvedPart)
		}
		seenPart[edge.ResolvedPart] = true
		out = append(out, edge.ResolvedPart)
	}
	if len(out) != len(edges) {
		return nil, fmt.Errorf("orphan slide relationship")
	}
	return out, nil
}

func creationAttrs(node losslessxml.Element) map[xml.Name]string {
	out := map[xml.Name]string{}
	for _, attr := range node.Attributes() {
		out[attr.Name] = attr.Value
	}
	return out
}
func creationDirectChildren(doc *losslessxml.Document, root losslessxml.Element) []losslessxml.Element {
	out := []losslessxml.Element{}
	for _, node := range doc.Elements() {
		p, ok := node.Parent()
		if ok && p == root {
			out = append(out, node)
		}
	}
	return out
}
func creationInsertion(old, current []byte, root, child xml.Name, want map[xml.Name]string) error {
	a, err := losslessxml.Parse(old)
	if err != nil {
		return err
	}
	b, err := losslessxml.Parse(current)
	if err != nil {
		return err
	}
	if len(a.Elements()) == 0 || len(b.Elements()) == 0 || a.Elements()[0].Name() != root || b.Elements()[0].Name() != root {
		return fmt.Errorf("registry/list root changed: %s", root.Local)
	}
	before := creationDirectChildren(a, a.Elements()[0])
	after := creationDirectChildren(b, b.Elements()[0])
	if len(after) != len(before)+1 {
		return fmt.Errorf("expected one inserted %s, got %d -> %d", child.Local, len(before), len(after))
	}
	for i, node := range before {
		start, end := node.SourceRange()
		lo, hi := after[i].SourceRange()
		if node.Name() != after[i].Name() || !bytes.Equal(old[start:end], current[lo:hi]) {
			return fmt.Errorf("old registry child %d changed", i)
		}
	}
	added := after[len(before)]
	if added.Name() != child || !reflect.DeepEqual(creationAttrs(added), want) {
		return fmt.Errorf("wrong inserted child/attributes %v: %v", added.Name(), creationAttrs(added))
	}
	if len(creationDirectChildren(b, added)) != 0 {
		return fmt.Errorf("inserted registry child has nested elements")
	}
	_, end := a.Elements()[0].ContentRange()
	if end < 0 || !bytes.Equal(old[:end], current[:end]) || !bytes.HasSuffix(current, old[end:]) {
		return fmt.Errorf("outside insertion bytes changed")
	}
	return nil
}
func creationAppendInsertions(before, after map[string][]byte, prior, next packaging.Graph, oldSlide, newSlide string) error {
	oldOrder, e := creationSlideOrder(before, prior)
	if e != nil {
		return e
	}
	newOrder, e := creationSlideOrder(after, next)
	if e != nil {
		return e
	}
	if len(oldOrder) != 1 || len(newOrder) != 2 || newOrder[0] != oldSlide || newOrder[1] != newSlide {
		return fmt.Errorf("slide order changed: %q -> %q", oldOrder, newOrder)
	}
	oldRID, newRID := "", ""
	for _, edge := range prior.Edges {
		if edge.Source == packaging.PresentationPath && edge.ResolvedPart == oldSlide && edge.Type == packaging.RelTypeSlide {
			oldRID = edge.ID
		}
	}
	for _, edge := range next.Edges {
		if edge.Source == packaging.PresentationPath && edge.ResolvedPart == newSlide && edge.Type == packaging.RelTypeSlide {
			newRID = edge.ID
		}
	}
	if oldRID == "" || newRID == "" || oldRID == newRID {
		return fmt.Errorf("slide relationship identity")
	}
	oldDoc, e := losslessxml.Parse(before[packaging.PresentationPath])
	if e != nil {
		return e
	}
	var max uint64
	for _, node := range oldDoc.Elements() {
		if node.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldId"}) {
			continue
		}
		id := creationAttrs(node)[xml.Name{Local: "id"}]
		v, er := strconv.ParseUint(id, 10, 32)
		if er != nil {
			return er
		}
		if v > max {
			max = v
		}
	}
	if e = creationListInsertion(before[packaging.PresentationPath], after[packaging.PresentationPath], map[xml.Name]string{{Local: "id"}: strconv.FormatUint(max+1, 10), {Space: packaging.NSDocumentRelationships, Local: "id"}: newRID}); e != nil {
		return e
	}
	target := ""
	for _, edge := range next.Edges {
		if edge.Source == packaging.PresentationPath && edge.ID == newRID {
			target = edge.Target
		}
	}
	if target == "" {
		return fmt.Errorf("new slide relationship target empty")
	}
	if e = creationInsertion(before["ppt/_rels/presentation.xml.rels"], after["ppt/_rels/presentation.xml.rels"], xml.Name{Space: "http://schemas.openxmlformats.org/package/2006/relationships", Local: "Relationships"}, xml.Name{Space: "http://schemas.openxmlformats.org/package/2006/relationships", Local: "Relationship"}, map[xml.Name]string{{Local: "Id"}: newRID, {Local: "Type"}: packaging.RelTypeSlide, {Local: "Target"}: target}); e != nil {
		return e
	}
	if e = creationInsertionCollaterals(before["[Content_Types].xml"], after["[Content_Types].xml"], xml.Name{Space: "http://schemas.openxmlformats.org/package/2006/content-types", Local: "Types"}, xml.Name{Space: "http://schemas.openxmlformats.org/package/2006/content-types", Local: "Override"}, map[xml.Name]string{{Local: "PartName"}: "/" + newSlide, {Local: "ContentType"}: packaging.ContentTypeSlide}); e != nil {
		return e
	}
	// A reverse byte splice restores each complete old XML part, not just its
	// prefix/suffix. This bounds the insertion to the exact added element.
	for _, name := range []string{packaging.PresentationPath, "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"} {
		var partRoot, child xml.Name
		switch name {
		case packaging.PresentationPath:
			partRoot = xml.Name{Space: packaging.NSPresentationML, Local: "sldIdLst"}
			child = xml.Name{Space: packaging.NSPresentationML, Local: "sldId"}
		case "ppt/_rels/presentation.xml.rels":
			partRoot = xml.Name{Space: "http://schemas.openxmlformats.org/package/2006/relationships", Local: "Relationships"}
			child = xml.Name{Space: partRoot.Space, Local: "Relationship"}
		default:
			partRoot = xml.Name{Space: "http://schemas.openxmlformats.org/package/2006/content-types", Local: "Types"}
			child = xml.Name{Space: partRoot.Space, Local: "Override"}
		}
		if er := creationReverseInsertion(before[name], after[name], partRoot, child); er != nil {
			return fmt.Errorf("reverse insertion %s: %w", name, er)
		}
	}
	return nil
}
func creationListInsertion(old, current []byte, want map[xml.Name]string) error {
	a, e := losslessxml.Parse(old)
	if e != nil {
		return e
	}
	b, e := losslessxml.Parse(current)
	if e != nil {
		return e
	}
	list := xml.Name{Space: packaging.NSPresentationML, Local: "sldIdLst"}
	var before, after losslessxml.Element
	na, nb := 0, 0
	for _, node := range a.Elements() {
		if node.Name() == list {
			na++
			before = node
		}
	}
	for _, node := range b.Elements() {
		if node.Name() == list {
			nb++
			after = node
		}
	}
	if na != 1 || nb != 1 {
		return fmt.Errorf("slide list inventory %d/%d", na, nb)
	}
	prior := creationDirectChildren(a, before)
	updated := creationDirectChildren(b, after)
	if len(updated) != len(prior)+1 {
		return fmt.Errorf("slide-list child delta")
	}
	for i, node := range prior {
		start, end := node.SourceRange()
		lo, hi := updated[i].SourceRange()
		if node.Name() != updated[i].Name() || !bytes.Equal(old[start:end], current[lo:hi]) {
			return fmt.Errorf("old slide-list entry changed")
		}
	}
	added := updated[len(prior)]
	if added.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldId"}) || !reflect.DeepEqual(creationAttrs(added), want) || len(creationDirectChildren(b, added)) != 0 {
		return fmt.Errorf("unexpected slide-list insertion: %v %v", added.Name(), creationAttrs(added))
	}
	_, end := before.ContentRange()
	if end < 0 || !bytes.Equal(old[:end], current[:end]) || !bytes.HasSuffix(current, old[end:]) {
		return fmt.Errorf("slide list lexical prefix/suffix changed")
	}
	return nil
}

func creationReverseInsertion(old, current []byte, root, child xml.Name) error {
	a, e := losslessxml.Parse(old)
	if e != nil {
		return e
	}
	b, e := losslessxml.Parse(current)
	if e != nil {
		return e
	}
	var original, updated losslessxml.Element
	ca, cb := 0, 0
	for _, node := range a.Elements() {
		if node.Name() == root {
			ca++
			original = node
		}
	}
	for _, node := range b.Elements() {
		if node.Name() == root {
			cb++
			updated = node
		}
	}
	if ca != 1 || cb != 1 {
		return fmt.Errorf("owner count %d/%d", ca, cb)
	}
	oldChildren := creationDirectChildren(a, original)
	newChildren := creationDirectChildren(b, updated)
	if len(newChildren) != len(oldChildren)+1 || newChildren[len(newChildren)-1].Name() != child {
		return fmt.Errorf("not one inserted child")
	}
	start, end := newChildren[len(newChildren)-1].SourceRange()
	if start < 0 || end < start || end > len(current) {
		return fmt.Errorf("invalid inserted child range")
	}
	reversed := append(bytes.Clone(current[:start]), current[end:]...)
	if !bytes.Equal(reversed, old) {
		return fmt.Errorf("old member not recovered by deleting selected insertion")
	}
	return nil
}
func creationInsertionCollaterals(old, current []byte, root, child xml.Name, want map[xml.Name]string) error {
	if e := creationInsertion(old, current, root, child, want); e != nil {
		return e
	}
	// First prove that the independent oracle itself admits the one expected
	// insertion; the following alterations must then be rejected.
	// Well-formed extra Default and unknown XML are rejected by the oracle,
	// even when all old bytes and the expected child remain intact.
	for _, extra := range []string{`<Default Extension="foo" ContentType="application/xml"/>`, `<Unknown/>`} {
		at := bytes.LastIndex(current, []byte("</"))
		if at < 0 {
			return fmt.Errorf("missing root close")
		}
		corrupt := append(append(bytes.Clone(current[:at]), []byte(extra)...), current[at:]...)
		if _, e := losslessxml.Parse(corrupt); e != nil {
			return fmt.Errorf("collateral not valid XML: %w", e)
		}
		if e := creationInsertion(old, corrupt, root, child, want); e == nil {
			return fmt.Errorf("oracle admitted extra XML %s", strings.TrimSpace(extra))
		}
	}
	return nil
}
