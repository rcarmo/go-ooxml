package presentation

import (
	"errors"
	"fmt"
	"math"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type FreeformCommand struct {
	Op string `json:"op"`
	X  *int64 `json:"x,omitempty"`
	Y  *int64 `json:"y,omitempty"`
}
type FreeformPath struct {
	Width    int64             `json:"width"`
	Height   int64             `json:"height"`
	Commands []FreeformCommand `json:"commands"`
}
type FreeformOptions struct {
	Name      *string `json:"name,omitempty"`
	Fill      *string `json:"fill,omitempty"`
	LineColor *string `json:"lineColor,omitempty"`
	LineWidth *int64  `json:"lineWidth,omitempty"`
}
type FreeformReceipt struct {
	ShapeID  uint32          `json:"shapeId"`
	PartName string          `json:"partName"`
	Geometry PictureGeometry `json:"geometry"`
	Path     FreeformPath    `json:"path"`
}

func graphicsFreeformUnsupported(s string) error { return editRefusal("PPTX_FREEFORM_UNSUPPORTED", s) }
func (s *EditSession) AddFreeform(part string, g PictureGeometry, p FreeformPath, o FreeformOptions) (FreeformReceipt, error) {
	var none FreeformReceipt
	if p.Width < 1 || p.Width > math.MaxInt32 || p.Height < 1 || p.Height > math.MaxInt32 || len(p.Commands) < 2 || len(p.Commands) > 256 {
		return none, graphicsFreeformUnsupported("freeform dimensions/list bounds")
	}
	fill := "none"
	if o.Fill != nil {
		fill = *o.Fill
	}
	active, closed := false, false
	lines := 0
	vertices := map[[2]int64]bool{}
	last := [2]int64{}
	finish := func() bool { return !active || lines >= 1 && (fill == "none" || closed) }
	body := ""
	commands := []FreeformCommand{}
	for _, c := range p.Commands {
		if c.Op == "close" {
			if c.X != nil || c.Y != nil || !active || closed || lines < 2 || len(vertices) < 3 {
				return none, graphicsFreeformUnsupported("polygonal close required")
			}
			commands = append(commands, FreeformCommand{Op: "close"})
			closed = true
			body += `<a:close/>`
			continue
		}
		if c.Op != "move" && c.Op != "line" || c.X == nil || c.Y == nil || *c.X < 0 || *c.X > p.Width || *c.Y < 0 || *c.Y > p.Height {
			return none, graphicsFreeformUnsupported("freeform command/point")
		}
		point := [2]int64{*c.X, *c.Y}
		if c.Op == "move" {
			if !finish() {
				return none, graphicsFreeformUnsupported("incomplete subpath")
			}
			active, closed = true, false
			lines = 0
			vertices = map[[2]int64]bool{}
			last = point
			vertices[point] = true
		} else {
			if !active || closed || point == last {
				return none, graphicsFreeformUnsupported("line requires active unique point")
			}
			lines++
			last = point
			vertices[point] = true
		}
		x, y := *c.X, *c.Y
		commands = append(commands, FreeformCommand{c.Op, &x, &y})
		tag := "moveTo"
		if c.Op == "line" {
			tag = "lnTo"
		}
		body += fmt.Sprintf(`<a:%s><a:pt x="%d" y="%d"/></a:%s>`, tag, x, y, tag)
	}
	if !finish() {
		return none, graphicsFreeformUnsupported("incomplete final subpath")
	}
	if !s.slides[part] {
		return none, graphicsFreeformUnsupported("slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return none, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, h, e := s.pkg.Part(part)
	if e != nil {
		return none, e
	}
	next, preset, e := graphicsAutoShapeXML(source, part, "rect", g, AutoShapeOptions{Name: o.Name, Fill: &fill, LineColor: o.LineColor, LineWidth: o.LineWidth})
	if e != nil {
		var r *packaging.Refusal
		if errors.As(e, &r) && r.Kind != "PPTX_ID_EXHAUSTED" {
			e = graphicsFreeformUnsupported(r.Detail)
		}
		return none, e
	}
	doc, e := losslessxml.Parse(next)
	if e != nil {
		return none, e
	}
	var shape losslessxml.Element
	for _, n := range doc.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "sp") {
			continue
		}
		id, _, _, er := graphicsIdentity(doc, n, "nvSpPr")
		if er != nil {
			return none, er
		}
		if id == preset.ShapeID {
			shape = n
		}
	}
	props, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
	if e != nil {
		return none, e
	}
	geometry, e := graphicsOne(doc, props, packaging.NSDrawingML, "prstGeom", false)
	if e != nil {
		return none, e
	}
	mode := "norm"
	if fill == "none" {
		mode = "none"
	}
	custom := fmt.Sprintf(`<a:custGeom xmlns:a="%s"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="%d" h="%d" fill="%s" stroke="1" extrusionOk="0">%s</a:path></a:pathLst></a:custGeom>`, packaging.NSDrawingML, p.Width, p.Height, mode, body)
	a, z := geometry.SourceRange()
	next, e = fmtSplice(doc, next, a, z, []byte(custom))
	if e != nil {
		return none, e
	}
	doc, e = losslessxml.Parse(next)
	if e != nil {
		return none, e
	}
	for _, n := range doc.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "sp") {
			continue
		}
		id, _, _, er := graphicsIdentity(doc, n, "nvSpPr")
		if er != nil {
			return none, er
		}
		if id == preset.ShapeID {
			shape = n
		}
	}
	textBody, e := graphicsOne(doc, shape, packaging.NSPresentationML, "txBody", false)
	if e != nil {
		return none, e
	}
	next, e = doc.RemoveElements([]losslessxml.Element{textBody})
	if e != nil {
		return none, e
	}
	if o.Name == nil {
		doc, e = losslessxml.Parse(next)
		if e != nil {
			return none, e
		}
		for _, n := range doc.Elements() {
			if n.Name() != name(packaging.NSPresentationML, "cNvPr") {
				continue
			}
			raw, _ := graphicsAttr(n, "", "id")
			id, _ := graphicsNumber(raw, 1, math.MaxInt32)
			if uint32(id) == preset.ShapeID {
				next, e = doc.Edit(nil, []losslessxml.AttributeEdit{{Target: n, Name: name("", "name"), Value: fmt.Sprintf("Freeform %d", id)}})
				if e != nil {
					return none, e
				}
				break
			}
		}
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: next}}); e != nil {
		return none, e
	}
	s.generation++
	return FreeformReceipt{preset.ShapeID, part, g, FreeformPath{p.Width, p.Height, commands}}, nil
}
