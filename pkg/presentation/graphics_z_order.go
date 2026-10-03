package presentation

import (
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"math"
	"strings"
)

func graphicsOrderUnsupported(detail string) error {
	return editRefusal("PPTX_Z_ORDER_UNSUPPORTED", detail)
}
func (s *EditSession) graphicsOrderSelection(part string, groupID *uint32) ([]byte, string, []losslessxml.Element, []uint32, error) {
	if !s.slides[part] {
		return nil, "", nil, nil, graphicsOrderUnsupported("slide not enrolled")
	}
	b, h, e := s.pkg.Part(part)
	if e != nil {
		return nil, "", nil, nil, e
	}
	doc, e := losslessxml.Parse(b)
	if e != nil {
		return nil, "", nil, nil, e
	}
	es := doc.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, "", nil, nil, graphicsOrderUnsupported("slide root")
	}
	common, e := graphicsOne(doc, es[0], packaging.NSPresentationML, "cSld", false)
	if e != nil {
		return nil, "", nil, nil, graphicsOrderUnsupported("slide tree")
	}
	tree, e := graphicsOne(doc, common, packaging.NSPresentationML, "spTree", false)
	if e != nil {
		return nil, "", nil, nil, graphicsOrderUnsupported("slide tree")
	}
	seen := map[int64]bool{}
	for _, n := range es {
		if n.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		raw, _ := graphicsAttr(n, "", "id")
		v, er := graphicsNumber(raw, 1, math.MaxInt32)
		if er != nil || strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") || seen[v] {
			return nil, "", nil, nil, graphicsOrderUnsupported("duplicate/invalid identity")
		}
		seen[v] = true
	}
	parent := tree
	if groupID != nil {
		if *groupID == 0 || *groupID > math.MaxInt32 {
			return nil, "", nil, nil, graphicsOrderUnsupported("group identity")
		}
		count := 0
		for _, n := range es {
			if n.Name() != name(packaging.NSPresentationML, "grpSp") {
				continue
			}
			id, _, _, er := graphicsIdentity(doc, n, "nvGrpSpPr")
			if er != nil {
				return nil, "", nil, nil, graphicsOrderUnsupported("group metadata")
			}
			if id == *groupID {
				parent = n
				count++
			}
		}
		if count != 1 {
			return nil, "", nil, nil, graphicsOrderUnsupported("unique group parent")
		}
		for p, ok := parent.Parent(); ; p, ok = p.Parent() {
			if !ok {
				return nil, "", nil, nil, graphicsOrderUnsupported("group ancestry")
			}
			if p == tree {
				break
			}
			if p.Name() != name(packaging.NSPresentationML, "grpSp") {
				return nil, "", nil, nil, graphicsOrderUnsupported("group ancestry")
			}
		}
	}
	children := graphicsChildren(doc, parent)
	if len(children) < 2 || children[0].Name() != name(packaging.NSPresentationML, "nvGrpSpPr") || children[1].Name() != name(packaging.NSPresentationML, "grpSpPr") {
		return nil, "", nil, nil, graphicsOrderUnsupported("leading parent metadata")
	}
	if _, e = graphicsOne(doc, children[0], packaging.NSPresentationML, "cNvPr", false); e != nil {
		return nil, "", nil, nil, graphicsOrderUnsupported("parent identity")
	}
	if _, e = graphicsOne(doc, parent, packaging.NSPresentationML, "grpSpPr", false); e != nil {
		return nil, "", nil, nil, graphicsOrderUnsupported("parent transform")
	}
	owners := map[string]string{"sp": "nvSpPr", "pic": "nvPicPr", "grpSp": "nvGrpSpPr", "graphicFrame": "nvGraphicFramePr", "cxnSp": "nvCxnSpPr"}
	nodes := []losslessxml.Element{}
	ids := []uint32{}
	cursor, end := parent.ContentRange()
	for i, n := range children {
		a, z := n.SourceRange()
		if strings.TrimSpace(string(b[cursor:a])) != "" || n.Name().Space != packaging.NSPresentationML {
			return nil, "", nil, nil, graphicsOrderUnsupported("lexical/foreign sibling")
		}
		cursor = z
		if i < 2 {
			continue
		}
		if n.Name().Local == "extLst" {
			if i != len(children)-1 {
				return nil, "", nil, nil, graphicsOrderUnsupported("nonterminal extension")
			}
			continue
		}
		owner := owners[n.Name().Local]
		if owner == "" {
			return nil, "", nil, nil, graphicsOrderUnsupported("unsupported graphical sibling")
		}
		id, _, _, er := graphicsIdentity(doc, n, owner)
		if er != nil {
			return nil, "", nil, nil, graphicsOrderUnsupported("graphical identity")
		}
		nodes = append(nodes, n)
		ids = append(ids, id)
	}
	if strings.TrimSpace(string(b[cursor:end])) != "" {
		return nil, "", nil, nil, graphicsOrderUnsupported("lexical sibling barrier")
	}
	return b, h, nodes, ids, nil
}
func (s *EditSession) GetShapeOrder(part string, groupID *uint32) ([]uint32, error) {
	_, _, _, ids, e := s.graphicsOrderSelection(part, groupID)
	return ids, e
}
func (s *EditSession) ReorderShapes(part string, order []uint32, groupID *uint32) (int, error) {
	if len(order) > 1000 {
		return 0, graphicsOrderUnsupported("bounded order list")
	}
	selected := map[uint32]bool{}
	for _, id := range order {
		if id == 0 || id > math.MaxInt32 || selected[id] {
			return 0, graphicsOrderUnsupported("unique bounded IDs")
		}
		selected[id] = true
	}
	if e := s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	b, h, nodes, ids, e := s.graphicsOrderSelection(part, groupID)
	if e != nil {
		return 0, e
	}
	lookup := map[uint32]int{}
	slots := []int{}
	for i, id := range ids {
		lookup[id] = i
		if selected[id] {
			slots = append(slots, i)
		}
	}
	for _, id := range order {
		if _, ok := lookup[id]; !ok {
			return 0, graphicsOrderUnsupported("requested non-direct identity")
		}
	}
	changed := false
	for i, slot := range slots {
		if ids[slot] != order[i] {
			changed = true
		}
	}
	if !changed {
		return 0, nil
	}
	next := append([]byte{}, b...)
	for i := len(slots) - 1; i >= 0; i-- {
		slot := slots[i]
		a, z := nodes[slot].SourceRange()
		raw := nodes[lookup[order[i]]].Raw()
		next = append(append(append([]byte{}, next[:a]...), raw...), next[z:]...)
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: next}}); e != nil {
		return 0, e
	}
	s.generation++
	return 1, nil
}
