package presentation

import (
	"errors"
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type ShapeGroupOptions struct {
	Name *string `json:"name,omitempty"`
}
type ShapeGroupReceipt struct {
	ShapeID  uint32          `json:"shapeId"`
	PartName string          `json:"partName"`
	ChildIDs []uint32        `json:"childIds"`
	Geometry PictureGeometry `json:"geometry"`
}

func graphicsGroupUnsupported(detail string) error {
	return editRefusal("PPTX_GROUP_UNSUPPORTED", detail)
}

// GroupShapes wraps a contiguous direct selection with identity coordinate mapping.
func (s *EditSession) GroupShapes(part string, ids []uint32, g PictureGeometry, o ShapeGroupOptions) (ShapeGroupReceipt, error) {
	var none ShapeGroupReceipt
	if len(ids) < 2 || len(ids) > 100 {
		return none, graphicsGroupUnsupported("2-100 shape IDs required")
	}
	selection := map[uint32]bool{}
	for _, id := range ids {
		if id == 0 || id > math.MaxInt32 || selection[id] {
			return none, graphicsGroupUnsupported("bounded unique selection")
		}
		selection[id] = true
	}
	if g.X < math.MinInt32 || g.X > math.MaxInt32 || g.Y < math.MinInt32 || g.Y > math.MaxInt32 || g.Width < 1 || g.Width > math.MaxInt32 || g.Height < 1 || g.Height > math.MaxInt32 {
		return none, graphicsGroupUnsupported("group rectangle bounds")
	}
	if o.Name != nil {
		if strings.TrimSpace(*o.Name) == "" {
			return none, graphicsGroupUnsupported("nonempty group name")
		}
		if _, e := graphicsXMLString(*o.Name); e != nil {
			return none, graphicsGroupUnsupported("invalid group name")
		}
	}
	if !s.slides[part] {
		return none, graphicsGroupUnsupported("slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return none, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, hash, e := s.pkg.Part(part)
	if e != nil {
		return none, e
	}
	id, _, e := graphicsShapeAppend(source)
	if e != nil {
		var r *packaging.Refusal
		if errors.As(e, &r) && r.Kind == "PPTX_PICTURE_UNSUPPORTED" {
			e = graphicsGroupUnsupported(r.Detail)
		}
		return none, e
	}
	doc, e := losslessxml.Parse(source)
	if e != nil {
		return none, e
	}
	common, e := graphicsOne(doc, doc.Elements()[0], packaging.NSPresentationML, "cSld", false)
	if e != nil {
		return none, graphicsGroupUnsupported("slide tree")
	}
	tree, e := graphicsOne(doc, common, packaging.NSPresentationML, "spTree", false)
	if e != nil {
		return none, graphicsGroupUnsupported("slide tree")
	}
	children := graphicsChildren(doc, tree)
	selected := []losslessxml.Element{}
	order := []uint32{}
	indices := []int{}
	for i, n := range children {
		nv := ""
		if n.Name() == name(packaging.NSPresentationML, "sp") {
			nv = "nvSpPr"
		} else if n.Name() == name(packaging.NSPresentationML, "pic") {
			nv = "nvPicPr"
		}
		if nv == "" {
			continue
		}
		childID, _, _, er := graphicsIdentity(doc, n, nv)
		if er != nil {
			return none, graphicsGroupUnsupported("shape identity")
		}
		if selection[childID] {
			selected = append(selected, n)
			order = append(order, childID)
			indices = append(indices, i)
		}
	}
	if len(selected) != len(ids) {
		return none, graphicsGroupUnsupported("direct shape/picture selection required")
	}
	for i, index := range indices {
		if index != indices[0]+i {
			return none, graphicsGroupUnsupported("contiguous selection required")
		}
	}
	for _, child := range selected {
		for _, n := range doc.Elements() {
			if !manipulationWithin(n, child) {
				continue
			}
			if n.Name() == name(packaging.NSPresentationML, "ph") {
				return none, graphicsGroupUnsupported("placeholder grouping")
			}
			if n.Name() == name(packaging.NSDrawingML, "spLocks") || n.Name() == name(packaging.NSDrawingML, "picLocks") {
				v, ok := graphicsAttr(n, "", "noGrp")
				if ok && v != "0" && v != "false" {
					return none, graphicsGroupUnsupported("grouping locked")
				}
			}
		}
	}
	for _, n := range doc.Elements() {
		if n.Name() == name(packaging.NSDrawingML, "stCxn") || n.Name() == name(packaging.NSDrawingML, "endCxn") {
			v, _ := graphicsAttr(n, "", "id")
			endpoint, _ := strconv.ParseUint(v, 10, 32)
			if selection[uint32(endpoint)] {
				return none, graphicsGroupUnsupported("attached endpoint grouping")
			}
		}
	}
	prefixes := map[string]bool{}
	for _, n := range doc.Elements() {
		for prefix := range n.Namespaces() {
			prefixes[prefix] = true
		}
		q := n.QualifiedName()
		if strings.Contains(q, ":") {
			prefixes[strings.SplitN(q, ":", 2)[0]] = true
		}
	}
	allocate := func(base string) string {
		p := base
		for i := 1; prefixes[p]; i++ {
			p = base + strconv.Itoa(i)
		}
		prefixes[p] = true
		return p
	}
	p, a := allocate("group"), allocate("draw")
	start, _ := selected[0].SourceRange()
	_, end := selected[len(selected)-1].SourceRange()
	label := fmt.Sprintf("Group %d", id)
	if o.Name != nil {
		label = *o.Name
	}
	escaped, _ := graphicsXMLString(label)
	wrapper := fmt.Sprintf(`<%s:grpSp xmlns:%s="%s" xmlns:%s="%s"><%s:nvGrpSpPr><%s:cNvPr id="%d" name="%s"/><%s:cNvGrpSpPr/><%s:nvPr/></%s:nvGrpSpPr><%s:grpSpPr><%s:xfrm><%s:off x="%d" y="%d"/><%s:ext cx="%d" cy="%d"/><%s:chOff x="%d" y="%d"/><%s:chExt cx="%d" cy="%d"/></%s:xfrm></%s:grpSpPr>%s</%s:grpSp>`, p, p, packaging.NSPresentationML, a, packaging.NSDrawingML, p, p, id, escaped, p, p, p, p, a, a, g.X, g.Y, a, g.Width, g.Height, a, g.X, g.Y, a, g.Width, g.Height, a, p, string(source[start:end]), p)
	next, e := fmtSplice(doc, source, start, end, []byte(wrapper))
	if e != nil {
		return none, graphicsGroupUnsupported("group wrapper XML")
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: next}}); e != nil {
		return none, e
	}
	s.generation++
	return ShapeGroupReceipt{id, part, order, g}, nil
}
