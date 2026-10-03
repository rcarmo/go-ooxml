package presentation

import (
	"math"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type GroupTransform = PictureGroupTransform
type GroupTransformPatch struct {
	X           *int64 `json:"x,omitempty"`
	Y           *int64 `json:"y,omitempty"`
	Width       *int64 `json:"width,omitempty"`
	Height      *int64 `json:"height,omitempty"`
	ChildX      *int64 `json:"childX,omitempty"`
	ChildY      *int64 `json:"childY,omitempty"`
	ChildWidth  *int64 `json:"childWidth,omitempty"`
	ChildHeight *int64 `json:"childHeight,omitempty"`
	Rotation    *int64 `json:"rotation,omitempty"`
	FlipH       *bool  `json:"flipH,omitempty"`
	FlipV       *bool  `json:"flipV,omitempty"`
}
type GroupPoint struct {
	X float64 `json:"x"`
	Y float64 `json:"y"`
}

func graphicsValidateFrame(t GroupTransform) error {
	for _, v := range []int64{t.X, t.Y, t.ChildX, t.ChildY} {
		if v < math.MinInt32 || v > math.MaxInt32 {
			return graphicsGroupUnsupported("group coordinate bounds")
		}
	}
	for _, v := range []int64{t.Width, t.Height, t.ChildWidth, t.ChildHeight} {
		if v < 1 || v > math.MaxInt32 {
			return graphicsGroupUnsupported("group positive extent bounds")
		}
	}
	if t.Rotation < 0 || t.Rotation > 21599999 {
		return graphicsGroupUnsupported("group rotation bounds")
	}
	return nil
}
func (s *EditSession) graphicsGroupSelection(part string, id uint32) (GroupTransform, *losslessxml.Document, losslessxml.Element, []byte, string, error) {
	var zero GroupTransform
	if !s.slides[part] || id == 0 || id > math.MaxInt32 {
		return zero, nil, losslessxml.Element{}, nil, "", graphicsGroupUnsupported("group/slide identity")
	}
	b, h, e := s.pkg.Part(part)
	if e != nil {
		return zero, nil, losslessxml.Element{}, nil, "", e
	}
	doc, e := losslessxml.Parse(b)
	if e != nil {
		return zero, nil, losslessxml.Element{}, nil, "", e
	}
	es := doc.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return zero, nil, losslessxml.Element{}, nil, "", graphicsGroupUnsupported("slide root")
	}
	common, e := graphicsOne(doc, es[0], packaging.NSPresentationML, "cSld", false)
	if e != nil {
		return zero, nil, losslessxml.Element{}, nil, "", graphicsGroupUnsupported("slide tree")
	}
	tree, e := graphicsOne(doc, common, packaging.NSPresentationML, "spTree", false)
	if e != nil {
		return zero, nil, losslessxml.Element{}, nil, "", graphicsGroupUnsupported("slide tree")
	}
	seen := map[int64]bool{}
	var group losslessxml.Element
	count := 0
	for _, n := range es {
		if n.Name() == name(packaging.NSPresentationML, "cNvPr") {
			raw, _ := graphicsAttr(n, "", "id")
			v, er := graphicsNumber(raw, 1, math.MaxInt32)
			if er != nil || seen[v] {
				return zero, nil, group, nil, "", graphicsGroupUnsupported("duplicate/invalid slide identity")
			}
			seen[v] = true
		}
		if n.Name() == name(packaging.NSPresentationML, "grpSp") {
			v, _, _, er := graphicsIdentity(doc, n, "nvGrpSpPr")
			if er != nil {
				return zero, nil, group, nil, "", graphicsGroupUnsupported("group identity grammar")
			}
			if v == id {
				group = n
				count++
			}
		}
	}
	if count != 1 {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("unique group required")
	}
	for p, ok := group.Parent(); ; p, ok = p.Parent() {
		if !ok {
			return zero, nil, group, nil, "", graphicsGroupUnsupported("group ancestry")
		}
		if p == tree {
			break
		}
		if p.Name() != name(packaging.NSPresentationML, "grpSp") {
			return zero, nil, group, nil, "", graphicsGroupUnsupported("group ancestry")
		}
	}
	props, e := graphicsOne(doc, group, packaging.NSPresentationML, "grpSpPr", false)
	if e != nil {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("group properties")
	}
	tr, e := graphicsOne(doc, props, packaging.NSDrawingML, "xfrm", false)
	if e != nil {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("group direct transform")
	}
	if e = graphicsAttributes(tr, []string{"rot", "flipH", "flipV"}, nil); e != nil {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("group transform attributes")
	}
	children := graphicsChildren(doc, tr)
	if len(children) != 4 {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("group transform children")
	}
	cursor, end := tr.ContentRange()
	for i, n := range children {
		if n.Name() != name(packaging.NSDrawingML, []string{"off", "ext", "chOff", "chExt"}[i]) {
			return zero, nil, group, nil, "", graphicsGroupUnsupported("group transform order")
		}
		a, z := n.SourceRange()
		c, d := n.ContentRange()
		if strings.TrimSpace(string(b[cursor:a])) != "" || len(graphicsChildren(doc, n)) != 0 || strings.TrimSpace(string(b[c:d])) != "" {
			return zero, nil, group, nil, "", graphicsGroupUnsupported("group lexical barrier")
		}
		keys := []string{"x", "y"}
		if i == 1 || i == 3 {
			keys = []string{"cx", "cy"}
		}
		if e = graphicsAttributes(n, keys, keys); e != nil {
			return zero, nil, group, nil, "", graphicsGroupUnsupported("group child attributes")
		}
		cursor = z
	}
	if strings.TrimSpace(string(b[cursor:end])) != "" {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("group lexical barrier")
	}
	value, e := graphicsReadTransform(doc, props, true)
	if e != nil || value == nil {
		return zero, nil, group, nil, "", graphicsGroupUnsupported("group scalar grammar")
	}
	if e = graphicsValidateFrame(*value); e != nil {
		return zero, nil, group, nil, "", e
	}
	return *value, doc, tr, b, h, nil
}
func (s *EditSession) GetGroupTransform(part string, id uint32) (GroupTransform, error) {
	v, _, _, _, _, e := s.graphicsGroupSelection(part, id)
	return v, e
}
func (s *EditSession) PatchGroupTransform(part string, id uint32, patch GroupTransformPatch) (int, error) {
	if e := s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	v, doc, tr, b, h, e := s.graphicsGroupSelection(part, id)
	if e != nil {
		return 0, e
	}
	next := v
	specs := []struct {
		old    *int64
		target *int64
		patch  *int64
		node   int
		key    string
	}{{&v.X, &next.X, patch.X, 0, "x"}, {&v.Y, &next.Y, patch.Y, 0, "y"}, {&v.Width, &next.Width, patch.Width, 1, "cx"}, {&v.Height, &next.Height, patch.Height, 1, "cy"}, {&v.ChildX, &next.ChildX, patch.ChildX, 2, "x"}, {&v.ChildY, &next.ChildY, patch.ChildY, 2, "y"}, {&v.ChildWidth, &next.ChildWidth, patch.ChildWidth, 3, "cx"}, {&v.ChildHeight, &next.ChildHeight, patch.ChildHeight, 3, "cy"}, {&v.Rotation, &next.Rotation, patch.Rotation, -1, "rot"}}
	for _, s := range specs {
		if s.patch != nil {
			*s.target = *s.patch
		}
	}
	if patch.FlipH != nil {
		next.FlipH = *patch.FlipH
	}
	if patch.FlipV != nil {
		next.FlipV = *patch.FlipV
	}
	if e = graphicsValidateFrame(next); e != nil {
		return 0, e
	}
	children := graphicsChildren(doc, tr)
	edits := []losslessxml.AttributeEdit{}
	for _, spec := range specs {
		if *spec.old == *spec.target {
			continue
		}
		node := tr
		if spec.node >= 0 {
			node = children[spec.node]
		}
		edits = append(edits, losslessxml.AttributeEdit{Target: node, Name: name("", spec.key), Value: strconv.FormatInt(*spec.target, 10)})
	}
	for _, spec := range []struct {
		k        string
		old, new bool
	}{{"flipH", v.FlipH, next.FlipH}, {"flipV", v.FlipV, next.FlipV}} {
		if spec.old != spec.new {
			value := "0"
			if spec.new {
				value = "1"
			}
			edits = append(edits, losslessxml.AttributeEdit{Target: tr, Name: name("", spec.k), Value: value})
		}
	}
	if len(edits) == 0 {
		return 0, nil
	}
	changed, e := doc.Edit(nil, edits)
	if e != nil {
		return 0, graphicsGroupUnsupported("group lexical edit")
	}
	return s.graphicsMetadataCommit(part, h, b, changed)
}
func graphicsPoint(p GroupPoint) (GroupPoint, error) {
	if math.IsNaN(p.X) || math.IsNaN(p.Y) || math.IsInf(p.X, 0) || math.IsInf(p.Y, 0) || math.Abs(p.X) > math.MaxInt32 || math.Abs(p.Y) > math.MaxInt32 {
		return GroupPoint{}, graphicsGroupUnsupported("finite bounded group point")
	}
	return p, nil
}
func graphicsTrig(rotation int64) (float64, float64) {
	if rotation%5400000 == 0 {
		switch rotation / 5400000 {
		case 0:
			return 1, 0
		case 1:
			return 0, 1
		case 2:
			return -1, 0
		case 3:
			return 0, -1
		}
	}
	radians := float64(rotation) / 60000 * math.Pi / 180
	return math.Cos(radians), math.Sin(radians)
}
func MapGroupPoint(t GroupTransform, input GroupPoint) (GroupPoint, error) {
	if e := graphicsValidateFrame(t); e != nil {
		return GroupPoint{}, e
	}
	p, e := graphicsPoint(input)
	if e != nil {
		return GroupPoint{}, e
	}
	centre, e := graphicsPoint(GroupPoint{float64(t.X) + float64(t.Width)/2, float64(t.Y) + float64(t.Height)/2})
	if e != nil {
		return GroupPoint{}, e
	}
	scaled, e := graphicsPoint(GroupPoint{float64(t.X) + (p.X-float64(t.ChildX))*float64(t.Width)/float64(t.ChildWidth), float64(t.Y) + (p.Y-float64(t.ChildY))*float64(t.Height)/float64(t.ChildHeight)})
	if e != nil {
		return GroupPoint{}, e
	}
	dx, dy := scaled.X-centre.X, scaled.Y-centre.Y
	if t.FlipH {
		dx = -dx
	}
	if t.FlipV {
		dy = -dy
	}
	reflected, e := graphicsPoint(GroupPoint{centre.X + dx, centre.Y + dy})
	if e != nil {
		return GroupPoint{}, e
	}
	c, s := graphicsTrig(t.Rotation)
	dx, dy = reflected.X-centre.X, reflected.Y-centre.Y
	return graphicsPoint(GroupPoint{centre.X + c*dx - s*dy, centre.Y + s*dx + c*dy})
}
func UnmapGroupPoint(t GroupTransform, input GroupPoint) (GroupPoint, error) {
	if e := graphicsValidateFrame(t); e != nil {
		return GroupPoint{}, e
	}
	p, e := graphicsPoint(input)
	if e != nil {
		return GroupPoint{}, e
	}
	centre, e := graphicsPoint(GroupPoint{float64(t.X) + float64(t.Width)/2, float64(t.Y) + float64(t.Height)/2})
	if e != nil {
		return GroupPoint{}, e
	}
	c, s := graphicsTrig(t.Rotation)
	dx, dy := p.X-centre.X, p.Y-centre.Y
	unrotated, e := graphicsPoint(GroupPoint{centre.X + c*dx + s*dy, centre.Y - s*dx + c*dy})
	if e != nil {
		return GroupPoint{}, e
	}
	x, y := unrotated.X-centre.X, unrotated.Y-centre.Y
	if t.FlipH {
		x = -x
	}
	if t.FlipV {
		y = -y
	}
	reflected, e := graphicsPoint(GroupPoint{centre.X + x, centre.Y + y})
	if e != nil {
		return GroupPoint{}, e
	}
	return graphicsPoint(GroupPoint{float64(t.ChildX) + (reflected.X-float64(t.X))*float64(t.ChildWidth)/float64(t.Width), float64(t.ChildY) + (reflected.Y-float64(t.Y))*float64(t.ChildHeight)/float64(t.Height)})
}
func MapGroupPointChain(chain []GroupTransform, p GroupPoint) (GroupPoint, error) {
	if len(chain) < 1 || len(chain) > 16 {
		return GroupPoint{}, graphicsGroupUnsupported("1-16 mapping frames required")
	}
	var e error
	for i := len(chain) - 1; i >= 0; i-- {
		p, e = MapGroupPoint(chain[i], p)
		if e != nil {
			return GroupPoint{}, e
		}
	}
	return p, nil
}
func UnmapGroupPointChain(chain []GroupTransform, p GroupPoint) (GroupPoint, error) {
	if len(chain) < 1 || len(chain) > 16 {
		return GroupPoint{}, graphicsGroupUnsupported("1-16 mapping frames required")
	}
	var e error
	for _, t := range chain {
		p, e = UnmapGroupPoint(t, p)
		if e != nil {
			return GroupPoint{}, e
		}
	}
	return p, nil
}
