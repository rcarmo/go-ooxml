package presentation

import (
	"fmt"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"math"
	"reflect"
	"strings"
)

type GradientColorTransform struct {
	Kind  string `json:"kind"`
	Value int64  `json:"value"`
}
type GradientColor struct {
	Kind       string                   `json:"kind"`
	Value      string                   `json:"value"`
	Transforms []GradientColorTransform `json:"transforms"`
}
type GradientStop struct {
	Position int64         `json:"position"`
	Color    GradientColor `json:"color"`
}
type LinearGradient struct {
	Angle  int64          `json:"angle"`
	Scaled bool           `json:"scaled"`
	Stops  []GradientStop `json:"stops"`
}

func graphicsGradientUnsupported(s string) error { return editRefusal("PPTX_GRADIENT_UNSUPPORTED", s) }
func graphicsChoice(value string, allowed []string) bool {
	for _, v := range allowed {
		if value == v {
			return true
		}
	}
	return false
}

var graphicsSchemeColors = []string{"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink", "bg1", "tx1", "bg2", "tx2", "phClr"}

func graphicsGradientColor(c GradientColor) (GradientColor, error) {
	if !(c.Kind == "srgb" && retainedColor(c.Value) || c.Kind == "scheme" && graphicsChoice(c.Value, graphicsSchemeColors)) {
		return c, graphicsGradientUnsupported("colour reference")
	}
	if len(c.Transforms) > 5 {
		return c, graphicsGradientUnsupported("colour transform bounds")
	}
	seen := map[string]bool{}
	rows := []GradientColorTransform{}
	for _, t := range c.Transforms {
		if !graphicsChoice(t.Kind, []string{"tint", "shade", "lumMod", "lumOff", "alpha"}) || seen[t.Kind] || t.Value < 0 || t.Value > 100000 {
			return c, graphicsGradientUnsupported("colour transform")
		}
		seen[t.Kind] = true
		rows = append(rows, t)
	}
	c.Transforms = rows
	return c, nil
}
func graphicsGradientRequest(g LinearGradient) (LinearGradient, error) {
	if g.Angle < 0 || g.Angle > 21599999 || len(g.Stops) < 2 || len(g.Stops) > 16 {
		return g, graphicsGradientUnsupported("gradient angle/stops bounds")
	}
	out := g
	out.Stops = []GradientStop{}
	previous := int64(-1)
	for _, s := range g.Stops {
		if s.Position < 0 || s.Position > 100000 || s.Position <= previous {
			return g, graphicsGradientUnsupported("strict ascending stops")
		}
		previous = s.Position
		c, e := graphicsGradientColor(s.Color)
		if e != nil {
			return g, e
		}
		out.Stops = append(out.Stops, GradientStop{s.Position, c})
	}
	if out.Stops[0].Position != 0 || out.Stops[len(out.Stops)-1].Position != 100000 {
		return g, graphicsGradientUnsupported("boundary stops")
	}
	return out, nil
}
func graphicsGaps(source []byte, doc *losslessxml.Document, n losslessxml.Element) error {
	cursor, end := n.ContentRange()
	for _, c := range graphicsChildren(doc, n) {
		a, z := c.SourceRange()
		if strings.TrimSpace(string(source[cursor:a])) != "" {
			return fmt.Errorf("lexical barrier")
		}
		cursor = z
	}
	if strings.TrimSpace(string(source[cursor:end])) != "" {
		return fmt.Errorf("lexical barrier")
	}
	return nil
}
func graphicsStyleShape(source []byte, id uint32, kinds []string) (*losslessxml.Document, losslessxml.Element, error) {
	if id == 0 || id > math.MaxInt32 {
		return nil, losslessxml.Element{}, fmt.Errorf("shape identity")
	}
	doc, e := losslessxml.Parse(source)
	if e != nil {
		return nil, losslessxml.Element{}, e
	}
	es := doc.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, losslessxml.Element{}, fmt.Errorf("slide root")
	}
	common, e := graphicsOne(doc, es[0], packaging.NSPresentationML, "cSld", false)
	if e != nil {
		return nil, losslessxml.Element{}, e
	}
	tree, e := graphicsOne(doc, common, packaging.NSPresentationML, "spTree", false)
	if e != nil {
		return nil, losslessxml.Element{}, e
	}
	seen := map[int64]bool{}
	for _, n := range es {
		if n.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		raw, _ := graphicsAttr(n, "", "id")
		v, e := graphicsNumber(raw, 1, math.MaxInt32)
		if e != nil || strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") || seen[v] {
			return nil, losslessxml.Element{}, fmt.Errorf("duplicate/invalid identity")
		}
		seen[v] = true
	}
	count := 0
	var selected losslessxml.Element
	for _, n := range graphicsChildren(doc, tree) {
		if n.Name().Space != packaging.NSPresentationML || !graphicsChoice(n.Name().Local, kinds) {
			continue
		}
		nv := "nvSpPr"
		if n.Name().Local == "cxnSp" {
			nv = "nvCxnSpPr"
		}
		got, _, _, e := graphicsIdentity(doc, n, nv)
		if e != nil {
			return nil, selected, e
		}
		if got == id {
			selected = n
			count++
		}
	}
	if count != 1 {
		return nil, selected, fmt.Errorf("unique direct shape required")
	}
	return doc, selected, nil
}
func graphicsReadColor(source []byte, doc *losslessxml.Document, n losslessxml.Element) (GradientColor, error) {
	if n.Name().Space != packaging.NSDrawingML || !graphicsChoice(n.Name().Local, []string{"srgbClr", "schemeClr"}) {
		return GradientColor{}, graphicsGradientUnsupported("colour class")
	}
	if e := graphicsAttributes(n, []string{"val"}, []string{"val"}); e != nil {
		return GradientColor{}, graphicsGradientUnsupported("colour attributes")
	}
	if e := graphicsGaps(source, doc, n); e != nil {
		return GradientColor{}, graphicsGradientUnsupported("colour gap")
	}
	value, _ := graphicsAttr(n, "", "val")
	kind := "srgb"
	if n.Name().Local == "schemeClr" {
		kind = "scheme"
	}
	c := GradientColor{kind, value, []GradientColorTransform{}}
	for _, child := range graphicsChildren(doc, n) {
		if child.Name().Space != packaging.NSDrawingML || len(graphicsChildren(doc, child)) != 0 {
			return c, graphicsGradientUnsupported("colour transform grammar")
		}
		if e := graphicsAttributes(child, []string{"val"}, []string{"val"}); e != nil {
			return c, graphicsGradientUnsupported("colour transform attributes")
		}
		if e := graphicsGaps(source, doc, child); e != nil {
			return c, graphicsGradientUnsupported("colour transform gap")
		}
		raw, _ := graphicsAttr(child, "", "val")
		v, e := graphicsNumber(raw, 0, 100000)
		if e != nil {
			return c, graphicsGradientUnsupported("colour transform scalar")
		}
		c.Transforms = append(c.Transforms, GradientColorTransform{child.Name().Local, v})
	}
	return graphicsGradientColor(c)
}
func graphicsReadGradient(source []byte, doc *losslessxml.Document, n losslessxml.Element) (LinearGradient, error) {
	var g LinearGradient
	if e := graphicsAttributes(n, []string{"rotWithShape"}, []string{"rotWithShape"}); e != nil {
		return g, graphicsGradientUnsupported("gradient attributes")
	}
	rotate, _ := graphicsAttr(n, "", "rotWithShape")
	flag, e := graphicsFlag(rotate)
	if e != nil || !flag {
		return g, graphicsGradientUnsupported("rotate-with-shape required")
	}
	if e = graphicsGaps(source, doc, n); e != nil {
		return g, graphicsGradientUnsupported("gradient gap")
	}
	children := graphicsChildren(doc, n)
	if len(children) != 2 || children[0].Name() != name(packaging.NSDrawingML, "gsLst") || children[1].Name() != name(packaging.NSDrawingML, "lin") {
		return g, graphicsGradientUnsupported("linear gradient only")
	}
	list, lin := children[0], children[1]
	if len(list.Attributes()) != 0 || graphicsGaps(source, doc, list) != nil || graphicsAttributes(lin, []string{"ang", "scaled"}, []string{"ang", "scaled"}) != nil || graphicsGaps(source, doc, lin) != nil || len(graphicsChildren(doc, lin)) != 0 {
		return g, graphicsGradientUnsupported("gradient child grammar")
	}
	raw, _ := graphicsAttr(lin, "", "ang")
	g.Angle, e = graphicsNumber(raw, 0, 21599999)
	if e != nil {
		return g, graphicsGradientUnsupported("gradient angle")
	}
	raw, _ = graphicsAttr(lin, "", "scaled")
	g.Scaled, e = graphicsFlag(raw)
	if e != nil {
		return g, graphicsGradientUnsupported("gradient scaling")
	}
	g.Stops = []GradientStop{}
	for _, stop := range graphicsChildren(doc, list) {
		if stop.Name() != name(packaging.NSDrawingML, "gs") || graphicsAttributes(stop, []string{"pos"}, []string{"pos"}) != nil || graphicsGaps(source, doc, stop) != nil {
			return g, graphicsGradientUnsupported("gradient stop grammar")
		}
		color := graphicsChildren(doc, stop)
		if len(color) != 1 {
			return g, graphicsGradientUnsupported("gradient stop colour")
		}
		raw, _ := graphicsAttr(stop, "", "pos")
		position, e := graphicsNumber(raw, 0, 100000)
		if e != nil {
			return g, graphicsGradientUnsupported("gradient stop position")
		}
		c, e := graphicsReadColor(source, doc, color[0])
		if e != nil {
			return g, e
		}
		g.Stops = append(g.Stops, GradientStop{position, c})
	}
	return graphicsGradientRequest(g)
}
func graphicsGradientSelection(source []byte, id uint32) (*losslessxml.Document, losslessxml.Element, losslessxml.Element, *LinearGradient, error) {
	doc, shape, e := graphicsStyleShape(source, id, []string{"sp"})
	if e != nil {
		return nil, shape, shape, nil, graphicsGradientUnsupported(e.Error())
	}
	props, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
	if e != nil {
		return nil, props, props, nil, graphicsGradientUnsupported("shape properties")
	}
	if graphicsGaps(source, doc, props) != nil {
		return nil, props, props, nil, graphicsGradientUnsupported("property gap")
	}
	order := []string{"xfrm", "prstGeom", "custGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"}
	last := -1
	var fill losslessxml.Element
	count := 0
	for _, n := range graphicsChildren(doc, props) {
		rank := -1
		for i, k := range order {
			if n.Name() == name(packaging.NSDrawingML, k) {
				rank = i
			}
		}
		if rank < 0 || rank <= last {
			return nil, props, fill, nil, graphicsGradientUnsupported("property order")
		}
		last = rank
		if graphicsChoice(n.Name().Local, []string{"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}) {
			fill = n
			count++
		}
	}
	if count > 1 {
		return nil, props, fill, nil, graphicsGradientUnsupported("multiple fills")
	}
	if count == 0 {
		return doc, props, fill, nil, nil
	}
	switch fill.Name().Local {
	case "gradFill":
		g, e := graphicsReadGradient(source, doc, fill)
		if e != nil {
			return nil, props, fill, nil, e
		}
		return doc, props, fill, &g, nil
	case "solidFill":
		children := graphicsChildren(doc, fill)
		if len(fill.Attributes()) != 0 || len(children) != 1 || graphicsGaps(source, doc, fill) != nil {
			return nil, props, fill, nil, graphicsGradientUnsupported("solid fill grammar")
		}
		if _, e := graphicsReadColor(source, doc, children[0]); e != nil {
			return nil, props, fill, nil, e
		}
	case "noFill":
		if len(fill.Attributes()) != 0 || len(graphicsChildren(doc, fill)) != 0 || graphicsGaps(source, doc, fill) != nil {
			return nil, props, fill, nil, graphicsGradientUnsupported("noFill grammar")
		}
	default:
		return nil, props, fill, nil, graphicsGradientUnsupported("source fill class")
	}
	return doc, props, fill, nil, nil
}
func (s *EditSession) GetLinearGradient(part string, id uint32) (*LinearGradient, error) {
	if !s.slides[part] {
		return nil, graphicsGradientUnsupported("slide not enrolled")
	}
	b, _, e := s.pkg.Part(part)
	if e != nil {
		return nil, e
	}
	_, _, _, v, e := graphicsGradientSelection(b, id)
	return v, e
}
func (s *EditSession) SetLinearGradient(part string, id uint32, input LinearGradient) (int, error) {
	g, e := graphicsGradientRequest(input)
	if e != nil {
		return 0, e
	}
	if !s.slides[part] {
		return 0, graphicsGradientUnsupported("slide not enrolled")
	}
	if e = s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, h, e := s.pkg.Part(part)
	if e != nil {
		return 0, e
	}
	doc, props, fill, old, e := graphicsGradientSelection(source, id)
	if e != nil {
		return 0, e
	}
	if old != nil && reflect.DeepEqual(*old, g) {
		return 0, nil
	}
	stops := ""
	for _, s := range g.Stops {
		tag := "srgbClr"
		if s.Color.Kind == "scheme" {
			tag = "schemeClr"
		}
		transforms := ""
		for _, t := range s.Color.Transforms {
			transforms += fmt.Sprintf(`<a:%s val="%d"/>`, t.Kind, t.Value)
		}
		stops += fmt.Sprintf(`<a:gs pos="%d"><a:%s val="%s">%s</a:%s></a:gs>`, s.Position, tag, s.Color.Value, transforms, tag)
	}
	scaled := "0"
	if g.Scaled {
		scaled = "1"
	}
	node := fmt.Sprintf(`<a:gradFill xmlns:a="%s" rotWithShape="1"><a:gsLst>%s</a:gsLst><a:lin ang="%d" scaled="%s"/></a:gradFill>`, packaging.NSDrawingML, stops, g.Angle, scaled)
	a, z := fill.SourceRange()
	if a < 0 {
		_, a = props.ContentRange()
		for _, n := range graphicsChildren(doc, props) {
			if graphicsChoice(n.Name().Local, []string{"ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"}) {
				a, _ = n.SourceRange()
				break
			}
		}
		z = a
		if len(props.Raw()) >= 2 && strings.HasSuffix(string(props.Raw()), "/>") {
			raw := props.Raw()
			q := props.QualifiedName()
			value := string(raw[:len(raw)-2]) + ">" + node + "</" + q + ">"
			a, z = props.SourceRange()
			node = value
		}
	}
	next, e := fmtSplice(doc, source, a, z, []byte(node))
	if e != nil {
		return 0, e
	}
	return s.graphicsMetadataCommit(part, h, source, next)
}
