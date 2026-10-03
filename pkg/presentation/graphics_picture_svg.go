package presentation

import (
	"bytes"
	"regexp"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

type SvgPictureReceipt struct {
	PictureReceipt
	SvgMediaPart      string `json:"svgMediaPart"`
	SvgRelationshipID string `json:"svgRelationshipId"`
}
type graphicsSVGPair struct {
	data                []byte
	media, relationship string
}

var graphicsSVGPaint = regexp.MustCompile(`^(?:none|#[0-9a-fA-F]{3,8}|[a-zA-Z]+)$`)
var graphicsSVGURL = regexp.MustCompile(`(?i)url\s*\(`)

func graphicsValidateSVG(data []byte) error {
	if len(data) == 0 || len(data) > 1024*1024 || !utf8.Valid(data) {
		return graphicsUnsupported("SVG bounded UTF-8 bytes required")
	}
	upper := bytes.ToUpper(data)
	if bytes.Contains(upper, []byte("<!DOCTYPE")) || bytes.Contains(data, []byte("<?")) {
		return graphicsUnsupported("SVG declarations/PI")
	}
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return graphicsUnsupported("invalid SVG XML")
	}
	nodes := doc.Elements()
	if len(nodes) == 0 || nodes[0].Name() != name("http://www.w3.org/2000/svg", "svg") {
		return graphicsUnsupported("SVG namespace root")
	}
	tags := map[string]bool{}
	for _, k := range []string{"svg", "g", "rect", "circle", "ellipse", "line", "polyline", "polygon", "path", "title", "desc"} {
		tags[k] = true
	}
	attrs := map[string]bool{}
	for _, k := range []string{"width", "height", "viewBox", "preserveAspectRatio", "x", "y", "x1", "y1", "x2", "y2", "cx", "cy", "r", "rx", "ry", "points", "d", "fill", "stroke", "stroke-width", "opacity", "fill-opacity", "stroke-opacity", "fill-rule", "stroke-linecap", "stroke-linejoin", "transform"} {
		attrs[k] = true
	}
	for _, n := range nodes {
		if n.Name().Space != "http://www.w3.org/2000/svg" || !tags[n.Name().Local] {
			return graphicsUnsupported("unsupported SVG element")
		}
		for _, a := range n.Attributes() {
			if !(a.Name.Space == "" && attrs[a.Name.Local] || a.Name == name("http://www.w3.org/XML/1998/namespace", "space")) || graphicsSVGURL.MatchString(a.Value) {
				return graphicsUnsupported("SVG attribute/reference")
			}
			if (a.Name.Local == "fill" || a.Name.Local == "stroke") && !graphicsSVGPaint.MatchString(a.Value) {
				return graphicsUnsupported("SVG plain colour required")
			}
		}
	}
	return nil
}

// AddSvgPicture pairs passive source SVG with a caller-provided raster, never renders it.
func (s *EditSession) AddSvgPicture(part string, svg, fallback []byte, geometry PictureGeometry, options PictureOptions) (SvgPictureReceipt, error) {
	if err := graphicsValidateSVG(svg); err != nil {
		return SvgPictureReceipt{}, err
	}
	pair := &graphicsSVGPair{data: bytes.Clone(svg)}
	receipt, err := s.addPicture(part, fallback, geometry, options, PictureCrop{}, pair)
	if err != nil {
		return SvgPictureReceipt{}, err
	}
	return SvgPictureReceipt{receipt, pair.media, pair.relationship}, nil
}
