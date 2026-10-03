package presentation

import (
	"fmt"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"strconv"
	"strings"
)

func graphicsOpacityUnsupported(detail string) error {
	return editRefusal("PPTX_OPACITY_UNSUPPORTED", detail)
}
func graphicsAlphaValue(v int64) error {
	if v < 0 || v > 100000 {
		return graphicsOpacityUnsupported("alpha [0,100000] required")
	}
	return nil
}
func graphicsAlphaEffect(source []byte, doc *losslessxml.Document, n losslessxml.Element, key string) (int64, error) {
	if len(graphicsChildren(doc, n)) != 0 || graphicsAttributes(n, []string{key}, []string{key}) != nil || graphicsGaps(source, doc, n) != nil {
		return 0, graphicsOpacityUnsupported("alpha effect grammar")
	}
	raw, _ := graphicsAttr(n, "", key)
	if strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") {
		return 0, graphicsOpacityUnsupported("alpha scalar")
	}
	v, e := graphicsNumber(raw, 0, 100000)
	if e != nil {
		return 0, graphicsOpacityUnsupported("alpha scalar")
	}
	return v, nil
}
func (s *EditSession) graphicsAlphaSelection(part string, id uint32, picture bool) ([]byte, string, *losslessxml.Document, losslessxml.Element, losslessxml.Element, int64, error) {
	var empty losslessxml.Element
	if !s.slides[part] {
		return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("slide not enrolled")
	}
	source, h, e := s.pkg.Part(part)
	if e != nil {
		return nil, "", nil, empty, empty, 0, e
	}
	var doc *losslessxml.Document
	var owner losslessxml.Element
	if !picture {
		d, _, fill, g, e := graphicsGradientSelection(source, id)
		if e != nil || g != nil {
			detail := "solid fill required"
			if e != nil {
				detail = e.Error()
			}
			return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported(detail)
		}
		if fill.Name() != name(packaging.NSDrawingML, "solidFill") {
			return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("direct solid fill required")
		}
		doc = d
		owner = graphicsChildren(doc, fill)[0]
	} else {
		_, d, node, _, _, e := s.graphicsPictureSelection(part, id)
		if e != nil {
			return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported(e.Error())
		}
		doc = d
		fill, e := graphicsOne(doc, node, packaging.NSPresentationML, "blipFill", false)
		if e != nil {
			return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("picture fill")
		}
		owner, e = graphicsOne(doc, fill, packaging.NSDrawingML, "blip", false)
		if e != nil {
			return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("picture blip")
		}
		if graphicsGaps(source, doc, owner) != nil {
			return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("picture effect gap")
		}
		order := []string{"alphaModFix", "biLevel", "blur", "clrChange", "clrRepl", "duotone", "fillOverlay", "grayscl", "hsl", "lum", "tint", "extLst"}
		last := -1
		for _, n := range graphicsChildren(doc, owner) {
			rank := -1
			for i, k := range order {
				if n.Name() == name(packaging.NSDrawingML, k) {
					rank = i
				}
			}
			if rank < 0 || rank <= last {
				return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("picture competing effect grammar")
			}
			last = rank
		}
	}
	tag, key, value := "alpha", "val", int64(100000)
	if picture {
		tag, key, value = "alphaModFix", "amt", 0
	}
	var effect losslessxml.Element
	count := 0
	for _, n := range graphicsChildren(doc, owner) {
		if n.Name() == name(packaging.NSDrawingML, tag) {
			effect = n
			count++
		}
	}
	if count > 1 {
		return nil, "", nil, empty, empty, 0, graphicsOpacityUnsupported("duplicate alpha")
	}
	if count == 1 {
		v, e := graphicsAlphaEffect(source, doc, effect, key)
		if e != nil {
			return nil, "", nil, empty, empty, 0, e
		}
		value = v
		if picture {
			value = 100000 - v
		}
	}
	return source, h, doc, owner, effect, value, nil
}
func (s *EditSession) GetShapeOpacity(part string, id uint32) (int64, error) {
	_, _, _, _, _, v, e := s.graphicsAlphaSelection(part, id, false)
	return v, e
}
func (s *EditSession) GetPictureTransparency(part string, id uint32) (int64, error) {
	_, _, _, _, _, v, e := s.graphicsAlphaSelection(part, id, true)
	return v, e
}
func (s *EditSession) graphicsSetAlpha(part string, id uint32, input int64, picture bool) (int, error) {
	if e := graphicsAlphaValue(input); e != nil {
		return 0, e
	}
	if e := s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, h, doc, owner, effect, current, e := s.graphicsAlphaSelection(part, id, picture)
	if e != nil {
		return 0, e
	}
	if current == input {
		return 0, nil
	}
	value := input
	key, tag := "val", "alpha"
	if picture {
		value, key, tag = 100000-input, "amt", "alphaModFix"
	}
	var next []byte
	if effect.Name().Local != "" {
		next, e = doc.Edit(nil, []losslessxml.AttributeEdit{{Target: effect, Name: name("", key), Value: strconv.FormatInt(value, 10)}})
	} else {
		child := fmt.Sprintf(`<a:%s xmlns:a="%s" %s="%d"/>`, tag, packaging.NSDrawingML, key, value)
		raw := owner.Raw()
		if strings.HasSuffix(string(raw), "/>") {
			a, z := owner.SourceRange()
			content := string(raw[:len(raw)-2]) + ">" + child + "</" + owner.QualifiedName() + ">"
			next, e = fmtSplice(doc, source, a, z, []byte(content))
		} else {
			_, at := owner.ContentRange()
			if picture {
				children := graphicsChildren(doc, owner)
				if len(children) > 0 {
					at, _ = children[0].SourceRange()
				}
			}
			next, e = fmtSplice(doc, source, at, at, []byte(child))
		}
	}
	if e != nil {
		return 0, e
	}
	return s.graphicsMetadataCommit(part, h, source, next)
}
func (s *EditSession) SetShapeOpacity(part string, id uint32, value int64) (int, error) {
	return s.graphicsSetAlpha(part, id, value, false)
}
func (s *EditSession) SetPictureTransparency(part string, id uint32, value int64) (int, error) {
	return s.graphicsSetAlpha(part, id, value, true)
}
