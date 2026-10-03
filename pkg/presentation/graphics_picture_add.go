package presentation

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type PictureGeometry struct {
	X      int64 `json:"x"`
	Y      int64 `json:"y"`
	Width  int64 `json:"width"`
	Height int64 `json:"height"`
}
type PictureOptions struct {
	ContentType string  `json:"contentType"`
	Name        *string `json:"name,omitempty"`
	Description *string `json:"description,omitempty"`
}
type PictureReceipt struct {
	ShapeID        uint32 `json:"shapeId"`
	PartName       string `json:"partName"`
	MediaPart      string `json:"mediaPart"`
	RelationshipID string `json:"relationshipId"`
}

func graphicsUnsupported(detail string) error { return editRefusal("PPTX_PICTURE_UNSUPPORTED", detail) }
func graphicsXMLString(s string) (string, error) {
	if !utf8.ValidString(s) {
		return "", graphicsUnsupported("invalid XML UTF-8")
	}
	for _, r := range s {
		if !(r == 9 || r == 10 || r == 13 || r >= 32 && r <= 0xD7FF || r >= 0xE000 && r <= 0xFFFD || r >= 0x10000 && r <= 0x10FFFF) {
			return "", graphicsUnsupported("invalid XML character")
		}
	}
	var b bytes.Buffer
	if err := xml.EscapeText(&b, []byte(s)); err != nil {
		return "", graphicsUnsupported("invalid XML string")
	}
	return b.String(), nil
}
func graphicsShapeAppend(source []byte) (uint32, int, error) {
	fail := func(reason string) (uint32, int, error) { return 0, 0, graphicsUnsupported(reason) }
	d, err := losslessxml.Parse(source)
	if err != nil {
		return 0, 0, err
	}
	es := d.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return fail("expected slide root")
	}
	common, err := graphicsOne(d, es[0], packaging.NSPresentationML, "cSld", false)
	if err != nil {
		return fail("unique cSld required")
	}
	tree, err := graphicsOne(d, common, packaging.NSPresentationML, "spTree", false)
	if err != nil {
		return fail("unique shape tree required")
	}
	children := graphicsChildren(d, tree)
	if len(children) < 2 || children[0].Name() != name(packaging.NSPresentationML, "nvGrpSpPr") || children[1].Name() != name(packaging.NSPresentationML, "grpSpPr") {
		return fail("leading group properties required")
	}
	cursor, end := tree.ContentRange()
	for _, c := range children {
		a, b := c.SourceRange()
		if strings.TrimSpace(string(source[cursor:a])) != "" {
			return fail("shape tree lexical barrier")
		}
		cursor = b
	}
	if strings.TrimSpace(string(source[cursor:end])) != "" {
		return fail("shape tree lexical barrier")
	}
	nv, err := graphicsOne(d, tree, packaging.NSPresentationML, "nvGrpSpPr", false)
	if err != nil {
		return fail("root group identity")
	}
	if _, err = graphicsOne(d, nv, packaging.NSPresentationML, "cNvPr", false); err != nil {
		return fail("root group identity")
	}
	props, err := graphicsOne(d, tree, packaging.NSPresentationML, "grpSpPr", false)
	if err != nil {
		return fail("root group properties")
	}
	valid := map[string]string{"sp": "nvSpPr", "grpSp": "nvGrpSpPr", "graphicFrame": "nvGraphicFramePr", "cxnSp": "nvCxnSpPr", "pic": "nvPicPr"}
	var validate func(losslessxml.Element) error
	validate = func(parent losslessxml.Element) error {
		for _, child := range graphicsChildren(d, parent) {
			if child.Name().Space != packaging.NSPresentationML {
				continue
			}
			if nvName, ok := valid[child.Name().Local]; ok {
				if _, _, _, e := graphicsIdentity(d, child, nvName); e != nil {
					return graphicsUnsupported("ambiguous shape identity")
				}
				if child.Name().Local == "grpSp" {
					if e := validate(child); e != nil {
						return e
					}
				}
			}
		}
		return nil
	}
	for i, c := range children {
		n := c.Name()
		if n.Space != packaging.NSPresentationML || !(n.Local == "nvGrpSpPr" || n.Local == "grpSpPr" || n.Local == "extLst" || valid[n.Local] != "") || n.Local == "extLst" && i != len(children)-1 {
			return fail("unsupported shape tree child")
		}
	}
	transforms := graphicsChildren(d, props)
	if len(transforms) > 1 {
		return fail("root group transform ambiguity")
	}
	if len(transforms) == 1 {
		tr := transforms[0]
		if tr.Name() != name(packaging.NSDrawingML, "xfrm") || len(tr.Attributes()) != 0 {
			return fail("root transform properties")
		}
		nodes := graphicsChildren(d, tr)
		if len(nodes) != 4 {
			return fail("root transform children")
		}
		for i, n := range nodes {
			if n.Name() != name(packaging.NSDrawingML, []string{"off", "ext", "chOff", "chExt"}[i]) || len(graphicsChildren(d, n)) != 0 || len(n.Attributes()) != 2 {
				return fail("root transform grammar")
			}
			keys := []string{"cx", "cy"}
			if i == 0 || i == 2 {
				keys = []string{"x", "y"}
			}
			for _, k := range keys {
				v, ok := graphicsAttr(n, "", k)
				if !ok || v == "" || strings.Trim(v, "0") != "" {
					return fail("nonidentity root transform")
				}
			}
		}
	}
	if err = validate(tree); err != nil {
		return 0, 0, err
	}
	seen := map[int64]bool{}
	max := int64(0)
	for _, n := range es {
		if n.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		raw, _ := graphicsAttr(n, "", "id")
		if strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") {
			return fail("invalid shape identity")
		}
		id, e := graphicsNumber(raw, 1, math.MaxInt32)
		if e != nil || seen[id] {
			return fail("invalid or duplicate shape identity")
		}
		seen[id] = true
		if id > max {
			max = id
		}
	}
	if max >= math.MaxInt32 {
		return 0, 0, editRefusal("PPTX_ID_EXHAUSTED", "no next shape ID")
	}
	at := end
	for _, c := range children {
		if c.Name() == name(packaging.NSPresentationML, "extLst") {
			at, _ = c.SourceRange()
		}
	}
	return uint32(max + 1), at, nil
}

// AddPicture appends explicit PNG/JPEG bytes in one validated graph commit.
func (s *EditSession) AddPicture(part string, payload []byte, g PictureGeometry, o PictureOptions) (PictureReceipt, error) {
	return s.addPicture(part, payload, g, o, PictureCrop{})
}
func (s *EditSession) addPicture(part string, payload []byte, g PictureGeometry, o PictureOptions, crop PictureCrop) (PictureReceipt, error) {
	var none PictureReceipt
	if len(payload) == 0 || len(payload) > 64*1024*1024 {
		return none, graphicsUnsupported("payload length")
	}
	png := len(payload) >= 8 && bytes.Equal(payload[:8], []byte{137, 80, 78, 71, 13, 10, 26, 10})
	jpeg := len(payload) >= 4 && payload[0] == 255 && payload[1] == 216 && payload[len(payload)-2] == 255 && payload[len(payload)-1] == 217
	if !(o.ContentType == "image/png" && png || o.ContentType == "image/jpeg" && jpeg) {
		return none, graphicsUnsupported("MIME/signature mismatch")
	}
	if g.X < math.MinInt32 || g.X > math.MaxInt32 || g.Y < math.MinInt32 || g.Y > math.MaxInt32 || g.Width < 1 || g.Width > math.MaxInt32 || g.Height < 1 || g.Height > math.MaxInt32 {
		return none, graphicsUnsupported("bounded rectangle required")
	}
	if o.Name != nil && strings.TrimSpace(*o.Name) == "" {
		return none, graphicsUnsupported("nonempty picture name")
	}
	for _, p := range []*string{o.Name, o.Description} {
		if p != nil {
			if _, e := graphicsXMLString(*p); e != nil {
				return none, e
			}
		}
	}
	if !s.slides[part] {
		return none, graphicsUnsupported("slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return none, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, hash, err := s.pkg.Part(part)
	if err != nil {
		return none, err
	}
	id, at, err := graphicsShapeAppend(source)
	if err != nil {
		return none, err
	}
	graph, err := s.pkg.Graph()
	if err != nil {
		return none, err
	}
	used := map[string]bool{}
	relIDs := map[string]bool{}
	for _, p := range graph.Parts {
		used[strings.ToLower(p.Name)] = true
	}
	for _, e := range graph.Edges {
		if e.Source == part {
			relIDs[e.ID] = true
		}
	}
	extension := "png"
	if o.ContentType == "image/jpeg" {
		extension = "jpeg"
	}
	media := ""
	for i := 1; i <= len(used)+1; i++ {
		candidate := fmt.Sprintf("ppt/media/image%d.%s", i, extension)
		if !used[strings.ToLower(candidate)] {
			media = candidate
			break
		}
	}
	if media == "" {
		return none, graphicsUnsupported("no media identity")
	}
	rid := ""
	for i := 1; i <= len(relIDs)+1; i++ {
		candidate := "rId" + strconv.Itoa(i)
		if !relIDs[candidate] {
			rid = candidate
			break
		}
	}
	label := fmt.Sprintf("Picture %d", id)
	if o.Name != nil {
		label = *o.Name
	}
	escaped, _ := graphicsXMLString(label)
	description := ""
	if o.Description != nil {
		v, _ := graphicsXMLString(*o.Description)
		description = ` descr="` + v + `"`
	}
	picture := fmt.Sprintf(`<p:pic xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:nvPicPr><p:cNvPr id="%d" name="%s"%s/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`, packaging.NSPresentationML, packaging.NSDrawingML, packaging.NSDocumentRelationships, id, escaped, description, rid, g.X, g.Y, g.Width, g.Height)
	if crop != (PictureCrop{}) {
		node := fmt.Sprintf(`<a:srcRect l="%d" t="%d" r="%d" b="%d"/>`, crop.Left, crop.Top, crop.Right, crop.Bottom)
		picture = strings.Replace(picture, "<a:stretch>", node+"<a:stretch>", 1)
	}
	next := append(append(append([]byte{}, source[:at]...), []byte(picture)...), source[at:]...)
	plan, err := s.pkg.PlanGraphMutation(packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: media, ContentType: o.ContentType, Data: payload}}, Relationships: []packaging.RelationshipAddition{{Source: part, ID: rid, Type: packaging.RelTypeImage, TargetPart: media}}, Replacements: []packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: next}}})
	if err != nil {
		return none, err
	}
	if err = s.pkg.ApplyGraphPlan(plan); err != nil {
		return none, err
	}
	s.generation++
	return PictureReceipt{id, part, media, rid}, nil
}
