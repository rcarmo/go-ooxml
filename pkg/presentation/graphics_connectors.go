package presentation

import (
	"errors"
	"fmt"
	"math"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type ConnectorEndpoint struct {
	ShapeID uint32 `json:"shapeId"`
	Site    int    `json:"site"`
}
type ConnectorOptions struct {
	Name  *string `json:"name,omitempty"`
	Color *string `json:"color,omitempty"`
	Width *int64  `json:"width,omitempty"`
}
type ConnectorPoint struct {
	ConnectorEndpoint
	X int64 `json:"x"`
	Y int64 `json:"y"`
}
type ConnectorGeometry struct {
	X      int64 `json:"x"`
	Y      int64 `json:"y"`
	Width  int64 `json:"width"`
	Height int64 `json:"height"`
	FlipH  bool  `json:"flipH"`
	FlipV  bool  `json:"flipV"`
}
type ConnectorReceipt struct {
	ShapeID  uint32            `json:"shapeId"`
	PartName string            `json:"partName"`
	Start    ConnectorPoint    `json:"start"`
	End      ConnectorPoint    `json:"end"`
	Geometry ConnectorGeometry `json:"geometry"`
}

func graphicsConnectorUnsupported(s string) error {
	return editRefusal("PPTX_CONNECTOR_UNSUPPORTED", s)
}
func graphicsConnectorSite(doc *losslessxml.Document, tree losslessxml.Element, source []byte, e ConnectorEndpoint) (ConnectorPoint, error) {
	var none ConnectorPoint
	if e.ShapeID == 0 || e.ShapeID > math.MaxInt32 || e.Site < 0 || e.Site > 3 {
		return none, graphicsConnectorUnsupported("bounded endpoint identity/site")
	}
	var shape losslessxml.Element
	count := 0
	for _, n := range graphicsChildren(doc, tree) {
		if n.Name() != name(packaging.NSPresentationML, "sp") {
			continue
		}
		id, _, _, er := graphicsIdentity(doc, n, "nvSpPr")
		if er != nil {
			return none, graphicsConnectorUnsupported("endpoint identity")
		}
		if id == e.ShapeID {
			shape = n
			count++
		}
	}
	if count != 1 {
		return none, graphicsConnectorUnsupported("direct ordinary rectangle endpoint")
	}
	for _, n := range doc.Elements() {
		if !manipulationWithin(n, shape) {
			continue
		}
		if n.Name() == name(packaging.NSPresentationML, "ph") {
			return none, graphicsConnectorUnsupported("placeholder endpoint")
		}
		if n.Name() == name(packaging.NSDrawingML, "spLocks") {
			v, ok := graphicsAttr(n, "", "noConnect")
			if ok && v != "0" && v != "false" {
				return none, graphicsConnectorUnsupported("connection locked")
			}
		}
	}
	props, er := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
	if er != nil {
		return none, graphicsConnectorUnsupported("endpoint properties")
	}
	tr, er := graphicsOne(doc, props, packaging.NSDrawingML, "xfrm", false)
	if er != nil {
		return none, graphicsConnectorUnsupported("endpoint direct placement")
	}
	preset, er := graphicsOne(doc, props, packaging.NSDrawingML, "prstGeom", false)
	if er != nil {
		return none, graphicsConnectorUnsupported("endpoint preset")
	}
	if er = graphicsAttributes(preset, []string{"prst"}, []string{"prst"}); er != nil {
		return none, graphicsConnectorUnsupported("endpoint preset attributes")
	}
	kind, _ := graphicsAttr(preset, "", "prst")
	av, er := graphicsOne(doc, preset, packaging.NSDrawingML, "avLst", false)
	text, _ := av.Text()
	if kind != "rect" || er != nil || len(graphicsChildren(doc, preset)) != 1 || len(graphicsChildren(doc, av)) != 0 || len(av.Attributes()) != 0 || strings.TrimSpace(text) != "" {
		return none, graphicsConnectorUnsupported("plain rectangle preset required")
	}
	if er = graphicsAttributes(tr, []string{"rot", "flipH", "flipV"}, nil); er != nil {
		return none, graphicsConnectorUnsupported("endpoint transform attributes")
	}
	rot, ok := graphicsAttr(tr, "", "rot")
	if ok {
		v, er := graphicsNumber(rot, 0, math.MaxInt32)
		if er != nil || v != 0 {
			return none, graphicsConnectorUnsupported("rotated endpoint")
		}
	}
	for _, k := range []string{"flipH", "flipV"} {
		v, ok := graphicsAttr(tr, "", k)
		if ok && v != "0" && v != "false" {
			return none, graphicsConnectorUnsupported("flipped endpoint")
		}
	}
	children := graphicsChildren(doc, tr)
	if len(children) != 2 || children[0].Name() != name(packaging.NSDrawingML, "off") || children[1].Name() != name(packaging.NSDrawingML, "ext") {
		return none, graphicsConnectorUnsupported("endpoint transform order")
	}
	cursor, end := tr.ContentRange()
	for i, n := range children {
		a, z := n.SourceRange()
		text, _ := n.Text()
		if strings.TrimSpace(string(source[cursor:a])) != "" || len(graphicsChildren(doc, n)) != 0 || strings.TrimSpace(text) != "" {
			return none, graphicsConnectorUnsupported("endpoint lexical barrier")
		}
		keys := []string{"x", "y"}
		if i == 1 {
			keys = []string{"cx", "cy"}
		}
		if er = graphicsAttributes(n, keys, keys); er != nil {
			return none, graphicsConnectorUnsupported("endpoint coordinate attributes")
		}
		cursor = z
	}
	if strings.TrimSpace(string(source[cursor:end])) != "" {
		return none, graphicsConnectorUnsupported("endpoint lexical barrier")
	}
	values := []int64{}
	for i, n := range children {
		keys := []string{"x", "y"}
		minimum := int64(math.MinInt32)
		if i == 1 {
			keys = []string{"cx", "cy"}
			minimum = 1
		}
		for _, k := range keys {
			raw, _ := graphicsAttr(n, "", k)
			v, er := graphicsNumber(raw, minimum, math.MaxInt32)
			if er != nil {
				return none, graphicsConnectorUnsupported("endpoint coordinate bounds")
			}
			values = append(values, v)
		}
	}
	x, y, w, h := values[0], values[1], values[2], values[3]
	points := [][2]int64{{x + w/2, y}, {x, y + h/2}, {x + w/2, y + h}, {x + w, y + h/2}}
	p := points[e.Site]
	if p[0] < math.MinInt32 || p[0] > math.MaxInt32 || p[1] < math.MinInt32 || p[1] > math.MaxInt32 {
		return none, graphicsConnectorUnsupported("endpoint site overflow")
	}
	return ConnectorPoint{e, p[0], p[1]}, nil
}
func graphicsAddConnector(source []byte, part string, start, end ConnectorEndpoint, o ConnectorOptions) ([]byte, ConnectorReceipt, error) {
	var none ConnectorReceipt
	if start.ShapeID == end.ShapeID {
		return nil, none, graphicsConnectorUnsupported("distinct endpoints required")
	}
	label := ""
	if o.Name != nil {
		if strings.TrimSpace(*o.Name) == "" {
			return nil, none, graphicsConnectorUnsupported("nonempty connector name")
		}
		var er error
		label, er = graphicsXMLString(*o.Name)
		if er != nil {
			return nil, none, graphicsConnectorUnsupported("connector XML name")
		}
	}
	colour := "000000"
	if o.Color != nil {
		colour = *o.Color
	}
	if !retainedColor(colour) {
		return nil, none, graphicsConnectorUnsupported("connector sRGB")
	}
	width := int64(12700)
	if o.Width != nil {
		width = *o.Width
	}
	if width < 1 || width > 20116800 {
		return nil, none, graphicsConnectorUnsupported("connector line width")
	}
	id, at, er := graphicsShapeAppend(source)
	if er != nil {
		var r *packaging.Refusal
		if errors.As(er, &r) && r.Kind == "PPTX_PICTURE_UNSUPPORTED" {
			er = graphicsConnectorUnsupported(r.Detail)
		}
		return nil, none, er
	}
	if o.Name == nil {
		label = fmt.Sprintf("Connector %d", id)
	}
	doc, er := losslessxml.Parse(source)
	if er != nil {
		return nil, none, er
	}
	common, er := graphicsOne(doc, doc.Elements()[0], packaging.NSPresentationML, "cSld", false)
	if er != nil {
		return nil, none, graphicsConnectorUnsupported("slide tree")
	}
	tree, er := graphicsOne(doc, common, packaging.NSPresentationML, "spTree", false)
	if er != nil {
		return nil, none, graphicsConnectorUnsupported("slide tree")
	}
	first, er := graphicsConnectorSite(doc, tree, source, start)
	if er != nil {
		return nil, none, er
	}
	last, er := graphicsConnectorSite(doc, tree, source, end)
	if er != nil {
		return nil, none, er
	}
	g := ConnectorGeometry{X: min(first.X, last.X), Y: min(first.Y, last.Y), Width: last.X - first.X, Height: last.Y - first.Y, FlipH: last.X < first.X, FlipV: last.Y < first.Y}
	if g.Width < 0 {
		g.Width = -g.Width
	}
	if g.Height < 0 {
		g.Height = -g.Height
	}
	if g.Width > math.MaxInt32 || g.Height > math.MaxInt32 || g.Width == 0 && g.Height == 0 {
		return nil, none, graphicsConnectorUnsupported("connector extent/coincident sites")
	}
	fh, fv := "0", "0"
	if g.FlipH {
		fh = "1"
	}
	if g.FlipV {
		fv = "1"
	}
	shape := fmt.Sprintf(`<p:cxnSp xmlns:p="%s" xmlns:a="%s"><p:nvCxnSpPr><p:cNvPr id="%d" name="%s"/><p:cNvCxnSpPr><a:stCxn id="%d" idx="%d"/><a:endCxn id="%d" idx="%d"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm flipH="%s" flipV="%s"><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`, packaging.NSPresentationML, packaging.NSDrawingML, id, label, start.ShapeID, start.Site, end.ShapeID, end.Site, fh, fv, g.X, g.Y, g.Width, g.Height, width, colour)
	next, er := fmtSplice(doc, source, at, at, []byte(shape))
	if er != nil {
		return nil, none, er
	}
	return next, ConnectorReceipt{id, part, first, last, g}, nil
}
func (s *EditSession) AddConnector(part string, start, end ConnectorEndpoint, options ConnectorOptions) (ConnectorReceipt, error) {
	if !s.slides[part] {
		return ConnectorReceipt{}, graphicsConnectorUnsupported("slide not enrolled")
	}
	if er := s.manipulationProtection(); er != nil {
		return ConnectorReceipt{}, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, h, er := s.pkg.Part(part)
	if er != nil {
		return ConnectorReceipt{}, er
	}
	next, receipt, er := graphicsAddConnector(source, part, start, end, options)
	if er != nil {
		return ConnectorReceipt{}, er
	}
	if er = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: next}}); er != nil {
		return ConnectorReceipt{}, er
	}
	s.generation++
	return receipt, nil
}
