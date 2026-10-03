package presentation

import (
	"math"
	"regexp"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// PictureTransform reports direct DrawingML coordinates, never inherited placement.
type PictureTransform struct {
	X        int64 `json:"x"`
	Y        int64 `json:"y"`
	Width    int64 `json:"width"`
	Height   int64 `json:"height"`
	Rotation int64 `json:"rotation"`
	FlipH    bool  `json:"flipH"`
	FlipV    bool  `json:"flipV"`
}
type PictureGroupTransform struct {
	PictureTransform
	ChildX      int64 `json:"childX"`
	ChildY      int64 `json:"childY"`
	ChildWidth  int64 `json:"childWidth"`
	ChildHeight int64 `json:"childHeight"`
}
type PictureGroup struct {
	ShapeID   uint32                 `json:"shapeId"`
	Name      string                 `json:"name"`
	Transform *PictureGroupTransform `json:"transform"`
}
type PictureEmbeddedAsset struct {
	RelationshipID string `json:"relationshipId"`
	Target         string `json:"target"`
	PartName       string `json:"partName"`
	ContentType    string `json:"contentType"`
	ByteLength     int    `json:"byteLength"`
}
type PictureLinkedAsset struct {
	RelationshipID string `json:"relationshipId"`
	Target         string `json:"target"`
}
type PictureCrop struct {
	Left   int64 `json:"left"`
	Top    int64 `json:"top"`
	Right  int64 `json:"right"`
	Bottom int64 `json:"bottom"`
}

// PictureInfo is detached source-order identity, asset and local placement data.
type PictureInfo struct {
	ShapeID     uint32                `json:"shapeId"`
	Name        string                `json:"name"`
	Description *string               `json:"description"`
	SlidePart   string                `json:"slidePart"`
	Groups      []PictureGroup        `json:"groups"`
	Embedded    *PictureEmbeddedAsset `json:"embedded"`
	Linked      *PictureLinkedAsset   `json:"linked"`
	Transform   *PictureTransform     `json:"transform"`
	Crop        PictureCrop           `json:"crop"`
}

const graphicsSafeInteger int64 = 9007199254740991

var graphicsInteger = regexp.MustCompile(`^[-+]?\d+$`)
var graphicsImageMIME = regexp.MustCompile(`^image/[^\s/]+$`)

func pictureInvalid(detail string) error { return editRefusal("PPTX_PICTURE_INVALID", detail) }
func graphicsAttr(e losslessxml.Element, namespace, local string) (string, bool) {
	for _, a := range e.Attributes() {
		if a.Name == name(namespace, local) {
			return a.Value, true
		}
	}
	return "", false
}
func graphicsChildren(d *losslessxml.Document, parent losslessxml.Element) []losslessxml.Element {
	out := []losslessxml.Element{}
	for _, e := range d.Elements() {
		if p, ok := e.Parent(); ok && p == parent {
			out = append(out, e)
		}
	}
	return out
}
func graphicsOne(d *losslessxml.Document, parent losslessxml.Element, namespace, local string, optional bool) (losslessxml.Element, error) {
	rows := fmtChildren(d, parent, name(namespace, local))
	if optional && len(rows) == 0 {
		return losslessxml.Element{}, nil
	}
	if len(rows) != 1 {
		return losslessxml.Element{}, pictureInvalid("expected unique " + local)
	}
	return rows[0], nil
}
func graphicsNumber(raw string, min, max int64) (int64, error) {
	if !graphicsInteger.MatchString(raw) {
		return 0, pictureInvalid("invalid integer")
	}
	n, err := strconv.ParseInt(raw, 10, 64)
	if err != nil || n < min || n > max {
		return 0, pictureInvalid("integer out of bounds")
	}
	return n, nil
}
func graphicsFlag(raw string) (bool, error) {
	switch raw {
	case "", "0", "false":
		return false, nil
	case "1", "true":
		return true, nil
	default:
		return false, pictureInvalid("invalid Boolean")
	}
}
func graphicsIdentity(d *losslessxml.Document, parent losslessxml.Element, nv string) (uint32, string, *string, error) {
	node, err := graphicsOne(d, parent, packaging.NSPresentationML, nv, false)
	if err != nil {
		return 0, "", nil, err
	}
	props, err := graphicsOne(d, node, packaging.NSPresentationML, "cNvPr", false)
	if err != nil {
		return 0, "", nil, err
	}
	raw, _ := graphicsAttr(props, "", "id")
	id, err := graphicsNumber(raw, 1, math.MaxInt32)
	if err != nil {
		return 0, "", nil, err
	}
	label, _ := graphicsAttr(props, "", "name")
	var description *string
	if v, ok := graphicsAttr(props, "", "descr"); ok {
		description = &v
	}
	return uint32(id), label, description, nil
}
func graphicsReadTransform(d *losslessxml.Document, parent losslessxml.Element, group bool) (*PictureGroupTransform, error) {
	x, err := graphicsOne(d, parent, packaging.NSDrawingML, "xfrm", true)
	if err != nil {
		return nil, err
	}
	if x.Name().Local == "" {
		return nil, nil
	}
	off, err := graphicsOne(d, x, packaging.NSDrawingML, "off", false)
	if err != nil {
		return nil, err
	}
	ext, err := graphicsOne(d, x, packaging.NSDrawingML, "ext", false)
	if err != nil {
		return nil, err
	}
	value := &PictureGroupTransform{}
	fields := []struct {
		node losslessxml.Element
		key  string
		min  int64
		out  *int64
	}{{off, "x", -graphicsSafeInteger, &value.X}, {off, "y", -graphicsSafeInteger, &value.Y}, {ext, "cx", 0, &value.Width}, {ext, "cy", 0, &value.Height}}
	if group {
		co, e := graphicsOne(d, x, packaging.NSDrawingML, "chOff", false)
		if e != nil {
			return nil, e
		}
		ce, e := graphicsOne(d, x, packaging.NSDrawingML, "chExt", false)
		if e != nil {
			return nil, e
		}
		fields = append(fields, struct {
			node losslessxml.Element
			key  string
			min  int64
			out  *int64
		}{co, "x", -graphicsSafeInteger, &value.ChildX}, struct {
			node losslessxml.Element
			key  string
			min  int64
			out  *int64
		}{co, "y", -graphicsSafeInteger, &value.ChildY}, struct {
			node losslessxml.Element
			key  string
			min  int64
			out  *int64
		}{ce, "cx", 0, &value.ChildWidth}, struct {
			node losslessxml.Element
			key  string
			min  int64
			out  *int64
		}{ce, "cy", 0, &value.ChildHeight})
	}
	for _, f := range fields {
		raw, _ := graphicsAttr(f.node, "", f.key)
		n, e := graphicsNumber(raw, f.min, graphicsSafeInteger)
		if e != nil {
			return nil, e
		}
		*f.out = n
	}
	rot, ok := graphicsAttr(x, "", "rot")
	if !ok {
		rot = "0"
	}
	value.Rotation, err = graphicsNumber(rot, math.MinInt32, math.MaxInt32)
	if err != nil {
		return nil, err
	}
	h, _ := graphicsAttr(x, "", "flipH")
	v, _ := graphicsAttr(x, "", "flipV")
	value.FlipH, err = graphicsFlag(h)
	if err != nil {
		return nil, err
	}
	value.FlipV, err = graphicsFlag(v)
	if err != nil {
		return nil, err
	}
	return value, nil
}
func graphicsReadCrop(d *losslessxml.Document, fill losslessxml.Element) (PictureCrop, error) {
	rect, err := graphicsOne(d, fill, packaging.NSDrawingML, "srcRect", true)
	if err != nil {
		return PictureCrop{}, err
	}
	value := PictureCrop{}
	for _, f := range []struct {
		key string
		out *int64
	}{{"l", &value.Left}, {"t", &value.Top}, {"r", &value.Right}, {"b", &value.Bottom}} {
		raw, ok := graphicsAttr(rect, "", f.key)
		if !ok {
			raw = "0"
		}
		n, e := graphicsNumber(raw, math.MinInt32, math.MaxInt32)
		if e != nil {
			return PictureCrop{}, e
		}
		*f.out = n
	}
	return value, nil
}

// InspectPictures never decodes or fetches images, rewrites XML, or dirties handles.
func (s *EditSession) InspectPictures(part string) ([]PictureInfo, error) {
	if !s.slides[part] {
		return nil, pictureInvalid("slide not enrolled")
	}
	source, _, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(source)
	if err != nil {
		return nil, err
	}
	es := d.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, pictureInvalid("expected slide root")
	}
	common, err := graphicsOne(d, es[0], packaging.NSPresentationML, "cSld", false)
	if err != nil {
		return nil, err
	}
	tree, err := graphicsOne(d, common, packaging.NSPresentationML, "spTree", false)
	if err != nil {
		return nil, err
	}
	seen := map[int64]bool{}
	for _, e := range es {
		if e.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		raw, _ := graphicsAttr(e, "", "id")
		id, er := graphicsNumber(raw, 1, math.MaxInt32)
		if er != nil {
			return nil, er
		}
		if seen[id] {
			return nil, pictureInvalid("duplicate shape identity")
		}
		seen[id] = true
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return nil, err
	}
	types := map[string]string{}
	for _, p := range graph.Parts {
		types[p.Name] = p.ContentType
	}
	relation := func(id string) (packaging.Edge, error) {
		var result packaging.Edge
		count := 0
		for _, edge := range graph.Edges {
			if edge.Source == part && edge.ID == id {
				result = edge
				count++
			}
		}
		if count != 1 || result.Type != packaging.RelTypeImage {
			return result, pictureInvalid("missing or wrong image relationship")
		}
		return result, nil
	}
	result := []PictureInfo{}
	visited := map[losslessxml.Element]bool{}
	var visit func(losslessxml.Element, []PictureGroup) error
	visit = func(parent losslessxml.Element, groups []PictureGroup) error {
		for _, node := range graphicsChildren(d, parent) {
			switch node.Name() {
			case name(packaging.NSPresentationML, "pic"):
				visited[node] = true
				id, label, descr, e := graphicsIdentity(d, node, "nvPicPr")
				if e != nil {
					return e
				}
				fill, e := graphicsOne(d, node, packaging.NSPresentationML, "blipFill", false)
				if e != nil {
					return e
				}
				props, e := graphicsOne(d, node, packaging.NSPresentationML, "spPr", false)
				if e != nil {
					return e
				}
				blip, e := graphicsOne(d, fill, packaging.NSDrawingML, "blip", false)
				if e != nil {
					return e
				}
				for _, a := range blip.Attributes() {
					if (a.Name.Local == "embed" || a.Name.Local == "link") && a.Name.Space != packaging.NSDocumentRelationships {
						return pictureInvalid("wrong image attribute namespace")
					}
				}
				embed, _ := graphicsAttr(blip, packaging.NSDocumentRelationships, "embed")
				link, _ := graphicsAttr(blip, packaging.NSDocumentRelationships, "link")
				if embed == "" && link == "" {
					return pictureInvalid("picture has no relationship")
				}
				entry := PictureInfo{ShapeID: id, Name: label, Description: descr, SlidePart: part, Groups: append([]PictureGroup{}, groups...)}
				for i, g := range entry.Groups {
					if g.Transform != nil {
						copy := *g.Transform
						entry.Groups[i].Transform = &copy
					}
				}
				if embed != "" {
					edge, e := relation(embed)
					if e != nil {
						return e
					}
					if edge.External || edge.ResolvedPart == "" {
						return pictureInvalid("embedded image is external")
					}
					payload, _, e := s.pkg.Part(edge.ResolvedPart)
					if e != nil {
						return e
					}
					mime := types[edge.ResolvedPart]
					if !graphicsImageMIME.MatchString(mime) {
						return pictureInvalid("non-image content type")
					}
					entry.Embedded = &PictureEmbeddedAsset{embed, edge.Target, edge.ResolvedPart, mime, len(payload)}
				}
				if link != "" {
					edge, e := relation(link)
					if e != nil {
						return e
					}
					if !edge.External {
						return pictureInvalid("linked image must be external")
					}
					entry.Linked = &PictureLinkedAsset{link, edge.Target}
				}
				tr, e := graphicsReadTransform(d, props, false)
				if e != nil {
					return e
				}
				if tr != nil {
					entry.Transform = &tr.PictureTransform
				}
				entry.Crop, e = graphicsReadCrop(d, fill)
				if e != nil {
					return e
				}
				result = append(result, entry)
			case name(packaging.NSPresentationML, "grpSp"):
				id, label, _, e := graphicsIdentity(d, node, "nvGrpSpPr")
				if e != nil {
					return e
				}
				props, e := graphicsOne(d, node, packaging.NSPresentationML, "grpSpPr", false)
				if e != nil {
					return e
				}
				tr, e := graphicsReadTransform(d, props, true)
				if e != nil {
					return e
				}
				ancestors := append(append([]PictureGroup{}, groups...), PictureGroup{id, label, tr})
				if e = visit(node, ancestors); e != nil {
					return e
				}
			}
		}
		return nil
	}
	if err = visit(tree, []PictureGroup{}); err != nil {
		return nil, err
	}
	for _, e := range es {
		if e.Name() == name(packaging.NSPresentationML, "pic") && !visited[e] {
			return nil, pictureInvalid("misplaced picture")
		}
	}
	return result, nil
}
