package presentation

import (
	"bytes"
	"math"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type PictureTransformPatch struct {
	Rotation *int64 `json:"rotation,omitempty"`
	FlipH    *bool  `json:"flipH,omitempty"`
	FlipV    *bool  `json:"flipV,omitempty"`
}

func graphicsBoundedCrop(c PictureCrop) error {
	if c.Left < 0 || c.Left > 99999 || c.Top < 0 || c.Top > 99999 || c.Right < 0 || c.Right > 99999 || c.Bottom < 0 || c.Bottom > 99999 || c.Left+c.Right >= 100000 || c.Top+c.Bottom >= 100000 {
		return graphicsUnsupported("bounded crop requires visible region")
	}
	return nil
}
func (s *EditSession) graphicsPictureSelection(part string, id uint32) (PictureInfo, *losslessxml.Document, losslessxml.Element, []byte, string, error) {
	var empty PictureInfo
	if id == 0 || id > math.MaxInt32 {
		return empty, nil, losslessxml.Element{}, nil, "", graphicsUnsupported("bounded picture identity")
	}
	pics, e := s.InspectPictures(part)
	if e != nil {
		return empty, nil, losslessxml.Element{}, nil, "", e
	}
	var info *PictureInfo
	for i := range pics {
		if pics[i].ShapeID == id {
			info = &pics[i]
		}
	}
	if info == nil {
		return empty, nil, losslessxml.Element{}, nil, "", editRefusal("PPTX_PICTURE_NOT_FOUND", "picture identity absent")
	}
	b, h, e := s.pkg.Part(part)
	if e != nil {
		return empty, nil, losslessxml.Element{}, nil, "", e
	}
	d, e := losslessxml.Parse(b)
	if e != nil {
		return empty, nil, losslessxml.Element{}, nil, "", e
	}
	for _, n := range d.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "pic") {
			continue
		}
		got, _, _, e := graphicsIdentity(d, n, "nvPicPr")
		if e != nil {
			return empty, nil, losslessxml.Element{}, nil, "", e
		}
		if got == id {
			return *info, d, n, b, h, nil
		}
	}
	return empty, nil, losslessxml.Element{}, nil, "", graphicsUnsupported("selected picture mismatch")
}
func graphicsAttributes(node losslessxml.Element, allowed []string, required []string) error {
	seen := map[string]bool{}
	for _, a := range node.Attributes() {
		valid := false
		for _, k := range allowed {
			if a.Name.Space == "" && a.Name.Local == k {
				valid = true
				break
			}
		}
		if !valid {
			return graphicsUnsupported("unknown or foreign attribute")
		}
		seen[a.Name.Local] = true
	}
	for _, k := range required {
		if !seen[k] {
			return graphicsUnsupported("missing direct attribute")
		}
	}
	return nil
}
func graphicsCropSelection(d *losslessxml.Document, picture losslessxml.Element, source []byte) (losslessxml.Element, losslessxml.Element, error) {
	fill, e := graphicsOne(d, picture, packaging.NSPresentationML, "blipFill", false)
	if e != nil {
		return fill, losslessxml.Element{}, e
	}
	last := -1
	seen := map[string]bool{}
	var rect losslessxml.Element
	for _, n := range graphicsChildren(d, fill) {
		rank := -1
		for i, k := range []string{"blip", "srcRect", "tile", "stretch", "extLst"} {
			if n.Name() == name(packaging.NSDrawingML, k) {
				rank = i
			}
		}
		if rank < 0 || rank <= last || seen[n.Name().Local] {
			return fill, rect, graphicsUnsupported("unsupported picture fill order")
		}
		last = rank
		seen[n.Name().Local] = true
		if n.Name().Local == "srcRect" {
			rect = n
		}
	}
	if seen["tile"] && seen["stretch"] {
		return fill, rect, graphicsUnsupported("ambiguous picture fill")
	}
	if rect.Name().Local != "" {
		if len(graphicsChildren(d, rect)) != 0 {
			return fill, rect, graphicsUnsupported("mixed crop node")
		}
		a, b := rect.ContentRange()
		if a >= 0 && b >= a && strings.TrimSpace(string(source[a:b])) != "" {
			return fill, rect, graphicsUnsupported("crop lexical barrier")
		}
		if e = graphicsAttributes(rect, []string{"l", "t", "r", "b"}, nil); e != nil {
			return fill, rect, e
		}
	}
	return fill, rect, nil
}
func (s *EditSession) GetPictureCrop(part string, id uint32) (PictureCrop, error) {
	info, d, picture, b, _, e := s.graphicsPictureSelection(part, id)
	if e != nil {
		return PictureCrop{}, e
	}
	if _, _, e = graphicsCropSelection(d, picture, b); e != nil {
		return PictureCrop{}, e
	}
	if e = graphicsBoundedCrop(info.Crop); e != nil {
		return PictureCrop{}, e
	}
	return info.Crop, nil
}
func (s *EditSession) graphicsMetadataCommit(part, hash string, b, next []byte) (int, error) {
	if bytes.Equal(b, next) {
		return 0, nil
	}
	if e := s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: next}}); e != nil {
		return 0, e
	}
	s.generation++
	return 1, nil
}
func (s *EditSession) SetPictureCrop(part string, id uint32, crop PictureCrop) (int, error) {
	if e := graphicsBoundedCrop(crop); e != nil {
		return 0, e
	}
	if e := s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	info, d, picture, b, h, e := s.graphicsPictureSelection(part, id)
	if e != nil {
		return 0, e
	}
	fill, rect, e := graphicsCropSelection(d, picture, b)
	if e != nil {
		return 0, e
	}
	if e = graphicsBoundedCrop(info.Crop); e != nil {
		return 0, e
	}
	if info.Crop == crop {
		return 0, nil
	}
	var next []byte
	if rect.Name().Local != "" {
		changes := []losslessxml.AttributeEdit{}
		for _, v := range []struct {
			k string
			n int64
		}{{"l", crop.Left}, {"t", crop.Top}, {"r", crop.Right}, {"b", crop.Bottom}} {
			changes = append(changes, losslessxml.AttributeEdit{Target: rect, Name: name("", v.k), Value: strconv.FormatInt(v.n, 10)})
		}
		next, e = d.Edit(nil, changes)
	} else {
		_, at := fill.ContentRange()
		for _, n := range graphicsChildren(d, fill) {
			if n.Name().Local != "blip" {
				at, _ = n.SourceRange()
				break
			}
		}
		node := `<a:srcRect xmlns:a="` + packaging.NSDrawingML + `" l="` + strconv.FormatInt(crop.Left, 10) + `" t="` + strconv.FormatInt(crop.Top, 10) + `" r="` + strconv.FormatInt(crop.Right, 10) + `" b="` + strconv.FormatInt(crop.Bottom, 10) + `"/>`
		next, e = fmtSplice(d, b, at, at, []byte(node))
	}
	if e != nil {
		return 0, graphicsUnsupported("crop lexical edit")
	}
	return s.graphicsMetadataCommit(part, h, b, next)
}
func (s *EditSession) PatchPictureTransform(part string, id uint32, patch PictureTransformPatch) (int, error) {
	if patch.Rotation != nil && (*patch.Rotation < 0 || *patch.Rotation > 21599999) {
		return 0, graphicsUnsupported("bounded rotation")
	}
	if e := s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	info, d, picture, b, h, e := s.graphicsPictureSelection(part, id)
	if e != nil {
		return 0, e
	}
	if info.Transform == nil || info.Transform.Rotation < 0 || info.Transform.Rotation > 21599999 {
		return 0, graphicsUnsupported("direct bounded transform required")
	}
	props, e := graphicsOne(d, picture, packaging.NSPresentationML, "spPr", false)
	if e != nil {
		return 0, e
	}
	tr, e := graphicsOne(d, props, packaging.NSDrawingML, "xfrm", false)
	if e != nil {
		return 0, e
	}
	if e = graphicsAttributes(tr, []string{"rot", "flipH", "flipV"}, nil); e != nil {
		return 0, e
	}
	children := graphicsChildren(d, tr)
	if len(children) != 2 || children[0].Name() != name(packaging.NSDrawingML, "off") || children[1].Name() != name(packaging.NSDrawingML, "ext") {
		return 0, graphicsUnsupported("direct transform child order")
	}
	cursor, end := tr.ContentRange()
	for i, n := range children {
		a, z := n.SourceRange()
		if strings.TrimSpace(string(b[cursor:a])) != "" || len(graphicsChildren(d, n)) != 0 {
			return 0, graphicsUnsupported("transform lexical barrier")
		}
		cStart, cEnd := n.ContentRange()
		if strings.TrimSpace(string(b[cStart:cEnd])) != "" {
			return 0, graphicsUnsupported("transform mixed leaf")
		}
		keys := []string{"x", "y"}
		if i == 1 {
			keys = []string{"cx", "cy"}
		}
		if e = graphicsAttributes(n, keys, keys); e != nil {
			return 0, e
		}
		cursor = z
	}
	if strings.TrimSpace(string(b[cursor:end])) != "" {
		return 0, graphicsUnsupported("transform lexical barrier")
	}
	changes := []losslessxml.AttributeEdit{}
	if patch.Rotation != nil && *patch.Rotation != info.Transform.Rotation {
		changes = append(changes, losslessxml.AttributeEdit{Target: tr, Name: name("", "rot"), Value: strconv.FormatInt(*patch.Rotation, 10)})
	}
	for _, v := range []struct {
		k       string
		want    *bool
		current bool
	}{{"flipH", patch.FlipH, info.Transform.FlipH}, {"flipV", patch.FlipV, info.Transform.FlipV}} {
		if v.want != nil && *v.want != v.current {
			value := "0"
			if *v.want {
				value = "1"
			}
			changes = append(changes, losslessxml.AttributeEdit{Target: tr, Name: name("", v.k), Value: value})
		}
	}
	if len(changes) == 0 {
		return 0, nil
	}
	next, e := d.Edit(nil, changes)
	if e != nil {
		return 0, graphicsUnsupported("orientation lexical edit")
	}
	return s.graphicsMetadataCommit(part, h, b, next)
}
