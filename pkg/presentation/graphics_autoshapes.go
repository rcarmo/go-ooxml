package presentation

import (
	"errors"
	"fmt"
	"math"
	"strings"
	"unicode/utf16"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type AutoShapeOptions struct {
	Name        *string          `json:"name,omitempty"`
	Text        *string          `json:"text,omitempty"`
	Adjustments map[string]int64 `json:"adjustments,omitempty"`
	Fill        *string          `json:"fill,omitempty"`
	LineColor   *string          `json:"lineColor,omitempty"`
	LineWidth   *int64           `json:"lineWidth,omitempty"`
}
type AutoShapeReceipt struct {
	ShapeID     uint32           `json:"shapeId"`
	PartName    string           `json:"partName"`
	Preset      string           `json:"preset"`
	Geometry    PictureGeometry  `json:"geometry"`
	Adjustments map[string]int64 `json:"adjustments"`
}

func graphicsAutoShapeUnsupported(s string) error {
	return editRefusal("PPTX_AUTOSHAPE_UNSUPPORTED", s)
}
func graphicsAutoShapeXML(source []byte, part, preset string, g PictureGeometry, o AutoShapeOptions) ([]byte, AutoShapeReceipt, error) {
	var none AutoShapeReceipt
	if preset != "rect" && preset != "ellipse" && preset != "triangle" && preset != "diamond" && preset != "roundRect" {
		return nil, none, graphicsAutoShapeUnsupported("preset allowlist")
	}
	if g.X < 0 || g.X > math.MaxInt32 || g.Y < 0 || g.Y > math.MaxInt32 || g.Width < 1 || g.Width > math.MaxInt32 || g.Height < 1 || g.Height > math.MaxInt32 {
		return nil, none, graphicsAutoShapeUnsupported("shape geometry bounds")
	}
	values := map[string]int64{}
	for k, v := range o.Adjustments {
		if preset != "roundRect" || k != "adj" || v < 0 || v > 50000 {
			return nil, none, graphicsAutoShapeUnsupported("preset adjustment")
		}
	}
	if preset == "roundRect" {
		v := int64(16667)
		if a, ok := o.Adjustments["adj"]; ok {
			v = a
		}
		values["adj"] = v
	}
	text := ""
	if o.Text != nil {
		text = *o.Text
	}
	if len(utf16.Encode([]rune(text))) > 4096 {
		return nil, none, graphicsAutoShapeUnsupported("shape label bounds")
	}
	if _, e := graphicsXMLString(text); e != nil {
		return nil, none, graphicsAutoShapeUnsupported("shape label XML")
	}
	if o.Name != nil {
		if strings.TrimSpace(*o.Name) == "" {
			return nil, none, graphicsAutoShapeUnsupported("shape name")
		}
		if _, e := graphicsXMLString(*o.Name); e != nil {
			return nil, none, graphicsAutoShapeUnsupported("shape name XML")
		}
	}
	fill, line := "F2F2F2", "336699"
	width := int64(12700)
	if o.Fill != nil {
		fill = *o.Fill
	}
	if o.LineColor != nil {
		line = *o.LineColor
	}
	if o.LineWidth != nil {
		width = *o.LineWidth
	}
	if !(fill == "none" || retainedColor(fill)) || !(line == "none" || retainedColor(line)) || width < 0 || width > 20116800 {
		return nil, none, graphicsAutoShapeUnsupported("shape style")
	}
	id, at, e := graphicsShapeAppend(source)
	if e != nil {
		var r *packaging.Refusal
		if errors.As(e, &r) && r.Kind != "PPTX_ID_EXHAUSTED" {
			e = graphicsAutoShapeUnsupported(r.Detail)
		}
		return nil, none, e
	}
	label := fmt.Sprintf("AutoShape %d", id)
	if o.Name != nil {
		label = *o.Name
	}
	label, _ = graphicsXMLString(label)
	paint := func(c string) string {
		if c == "none" {
			return `<a:noFill/>`
		}
		return `<a:solidFill><a:srgbClr val="` + c + `"/></a:solidFill>`
	}
	guides := ""
	if v, ok := values["adj"]; ok {
		guides = fmt.Sprintf(`<a:gd name="adj" fmla="val %d"/>`, v)
	}
	body := ""
	text = strings.ReplaceAll(strings.ReplaceAll(text, "\r\n", "\n"), "\r", "\n")
	for _, row := range strings.Split(text, "\n") {
		escaped, _ := graphicsXMLString(row)
		body += `<a:p><a:r><a:t xml:space="preserve">` + escaped + `</a:t></a:r></a:p>`
	}
	shape := fmt.Sprintf(`<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>%s<a:ln w="%d">%s</a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>`, packaging.NSPresentationML, packaging.NSDrawingML, id, label, g.X, g.Y, g.Width, g.Height, preset, guides, paint(fill), width, paint(line), body)
	next := append(append(append([]byte{}, source[:at]...), []byte(shape)...), source[at:]...)
	return next, AutoShapeReceipt{id, part, preset, g, values}, nil
}
func (s *EditSession) AddAutoShape(part, preset string, g PictureGeometry, o AutoShapeOptions) (AutoShapeReceipt, error) {
	if !s.slides[part] {
		return AutoShapeReceipt{}, graphicsAutoShapeUnsupported("slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return AutoShapeReceipt{}, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, h, e := s.pkg.Part(part)
	if e != nil {
		return AutoShapeReceipt{}, e
	}
	next, receipt, e := graphicsAutoShapeXML(source, part, preset, g, o)
	if e != nil {
		return AutoShapeReceipt{}, e
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: next}}); e != nil {
		return AutoShapeReceipt{}, e
	}
	s.generation++
	return receipt, nil
}
