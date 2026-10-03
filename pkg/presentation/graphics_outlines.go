package presentation

import (
	"bytes"
	"encoding/json"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type OutlineJoin struct {
	Type  string `json:"type"`
	Limit *int64 `json:"limit,omitempty"`
}
type OutlineEnd struct {
	Type   string `json:"type"`
	Width  string `json:"width"`
	Length string `json:"length"`
}
type OutlineStyle struct {
	Cap      string       `json:"cap"`
	Compound string       `json:"compound"`
	Join     *OutlineJoin `json:"join"`
	HeadEnd  *OutlineEnd  `json:"headEnd"`
	TailEnd  *OutlineEnd  `json:"tailEnd"`
}

// Raw optional fields distinguish omitted decorations from explicit null removal.
type OutlinePatch struct {
	Cap      json.RawMessage `json:"cap,omitempty"`
	Compound json.RawMessage `json:"compound,omitempty"`
	Join     json.RawMessage `json:"join,omitempty"`
	HeadEnd  json.RawMessage `json:"headEnd,omitempty"`
	TailEnd  json.RawMessage `json:"tailEnd,omitempty"`
}

func graphicsOutlineUnsupported(detail string) error {
	return editRefusal("PPTX_OUTLINE_UNSUPPORTED", detail)
}
func graphicsOutlineJoin(j *OutlineJoin) error {
	if j == nil {
		return nil
	}
	if j.Type == "miter" {
		if j.Limit == nil || *j.Limit < 0 || *j.Limit > 1000000 {
			return graphicsOutlineUnsupported("miter limit")
		}
		return nil
	}
	if j.Limit != nil || j.Type != "round" && j.Type != "bevel" {
		return graphicsOutlineUnsupported("join enum/attributes")
	}
	return nil
}
func graphicsOutlineEnd(e *OutlineEnd) error {
	if e == nil {
		return nil
	}
	if !graphicsChoice(e.Type, []string{"none", "triangle", "stealth", "diamond", "oval", "arrow"}) || !graphicsChoice(e.Width, []string{"sm", "med", "lg"}) || !graphicsChoice(e.Length, []string{"sm", "med", "lg"}) {
		return graphicsOutlineUnsupported("arrowhead enum")
	}
	return nil
}
func graphicsDecodeOutline(raw json.RawMessage, value any) error {
	d := json.NewDecoder(bytes.NewReader(raw))
	d.DisallowUnknownFields()
	if e := d.Decode(value); e != nil {
		return graphicsOutlineUnsupported("decoration decoding")
	}
	return nil
}
func graphicsOutlineSelection(source []byte, id uint32) (*losslessxml.Document, losslessxml.Element, map[string]losslessxml.Element, OutlineStyle, error) {
	var zero OutlineStyle
	doc, shape, e := graphicsStyleShape(source, id, []string{"sp", "cxnSp"})
	if e != nil {
		return nil, shape, nil, zero, graphicsOutlineUnsupported(e.Error())
	}
	props, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
	if e != nil {
		return nil, props, nil, zero, graphicsOutlineUnsupported("shape properties")
	}
	line, e := graphicsOne(doc, props, packaging.NSDrawingML, "ln", false)
	if e != nil {
		return nil, line, nil, zero, graphicsOutlineUnsupported("existing direct line required")
	}
	if graphicsAttributes(line, []string{"w", "cap", "cmpd", "algn"}, nil) != nil || graphicsGaps(source, doc, line) != nil {
		return nil, line, nil, zero, graphicsOutlineUnsupported("line attributes/gap")
	}
	if raw, ok := graphicsAttr(line, "", "w"); ok {
		if strings.HasPrefix(raw, "+") || strings.HasPrefix(raw, "-") {
			return nil, line, nil, zero, graphicsOutlineUnsupported("line width")
		}
		if _, e := graphicsNumber(raw, 0, 20116800); e != nil {
			return nil, line, nil, zero, graphicsOutlineUnsupported("line width")
		}
	}
	if align, ok := graphicsAttr(line, "", "algn"); ok && !graphicsChoice(align, []string{"ctr", "in"}) {
		return nil, line, nil, zero, graphicsOutlineUnsupported("line alignment")
	}
	nodes := map[string]losslessxml.Element{}
	last := -1
	for _, n := range graphicsChildren(doc, line) {
		rank := -1
		if graphicsChoice(n.Name().Local, []string{"noFill", "solidFill", "gradFill", "pattFill"}) {
			rank = 0
		} else if graphicsChoice(n.Name().Local, []string{"prstDash", "custDash"}) {
			rank = 1
		} else if graphicsChoice(n.Name().Local, []string{"round", "bevel", "miter"}) {
			rank = 2
			nodes["join"] = n
		} else if n.Name().Local == "headEnd" {
			rank = 3
			nodes["headEnd"] = n
		} else if n.Name().Local == "tailEnd" {
			rank = 4
			nodes["tailEnd"] = n
		} else if n.Name().Local == "extLst" {
			rank = 5
		}
		if n.Name().Space != packaging.NSDrawingML || rank < 0 || rank <= last {
			return nil, line, nil, zero, graphicsOutlineUnsupported("line grammar/order")
		}
		last = rank
	}
	cap, ok := graphicsAttr(line, "", "cap")
	if !ok {
		cap = "flat"
	}
	compound, ok := graphicsAttr(line, "", "cmpd")
	if !ok {
		compound = "sng"
	}
	if !graphicsChoice(cap, []string{"flat", "rnd", "sq"}) || !graphicsChoice(compound, []string{"sng", "dbl", "thickThin", "thinThick", "tri"}) {
		return nil, line, nil, zero, graphicsOutlineUnsupported("line scalar enum")
	}
	value := OutlineStyle{Cap: cap, Compound: compound}
	if j, ok := nodes["join"]; ok {
		if len(graphicsChildren(doc, j)) != 0 || graphicsGaps(source, doc, j) != nil {
			return nil, line, nil, zero, graphicsOutlineUnsupported("join content")
		}
		allowed, required := []string{}, []string{}
		join := &OutlineJoin{Type: j.Name().Local}
		if join.Type == "miter" {
			allowed, required = []string{"lim"}, []string{"lim"}
			raw, _ := graphicsAttr(j, "", "lim")
			v, e := graphicsNumber(raw, 0, 1000000)
			if e != nil {
				return nil, line, nil, zero, graphicsOutlineUnsupported("miter limit")
			}
			join.Limit = &v
		}
		if graphicsAttributes(j, allowed, required) != nil || graphicsOutlineJoin(join) != nil {
			return nil, line, nil, zero, graphicsOutlineUnsupported("join attributes")
		}
		value.Join = join
	}
	for _, spec := range []struct {
		key    string
		target **OutlineEnd
	}{{"headEnd", &value.HeadEnd}, {"tailEnd", &value.TailEnd}} {
		n, ok := nodes[spec.key]
		if !ok {
			continue
		}
		if graphicsAttributes(n, []string{"type", "w", "len"}, nil) != nil || graphicsGaps(source, doc, n) != nil || len(graphicsChildren(doc, n)) != 0 {
			return nil, line, nil, zero, graphicsOutlineUnsupported("arrow content/attrs")
		}
		typ, ok := graphicsAttr(n, "", "type")
		if !ok {
			typ = "none"
		}
		w, ok := graphicsAttr(n, "", "w")
		if !ok {
			w = "med"
		}
		l, ok := graphicsAttr(n, "", "len")
		if !ok {
			l = "med"
		}
		end := &OutlineEnd{typ, w, l}
		if e = graphicsOutlineEnd(end); e != nil {
			return nil, line, nil, zero, e
		}
		*spec.target = end
	}
	return doc, line, nodes, value, nil
}
func (s *EditSession) GetOutlineStyle(part string, id uint32) (OutlineStyle, error) {
	if !s.slides[part] {
		return OutlineStyle{}, graphicsOutlineUnsupported("slide not enrolled")
	}
	source, _, e := s.pkg.Part(part)
	if e != nil {
		return OutlineStyle{}, e
	}
	_, _, _, value, e := graphicsOutlineSelection(source, id)
	return value, e
}
func (s *EditSession) PatchOutlineStyle(part string, id uint32, patch OutlinePatch) (int, error) {
	if !s.slides[part] {
		return 0, graphicsOutlineUnsupported("slide not enrolled")
	}
	if e := s.manipulationProtection(); e != nil {
		return 0, editRefusal("PPTX_PROTECTED", "presentation protection")
	}
	source, h, e := s.pkg.Part(part)
	if e != nil {
		return 0, e
	}
	doc, line, _, old, e := graphicsOutlineSelection(source, id)
	if e != nil {
		return 0, e
	}
	target := old
	if len(patch.Cap) > 0 {
		if e = graphicsDecodeOutline(patch.Cap, &target.Cap); e != nil || bytes.Equal(bytes.TrimSpace(patch.Cap), []byte("null")) || !graphicsChoice(target.Cap, []string{"flat", "rnd", "sq"}) {
			return 0, graphicsOutlineUnsupported("cap patch")
		}
	}
	if len(patch.Compound) > 0 {
		if e = graphicsDecodeOutline(patch.Compound, &target.Compound); e != nil || bytes.Equal(bytes.TrimSpace(patch.Compound), []byte("null")) || !graphicsChoice(target.Compound, []string{"sng", "dbl", "thickThin", "thinThick", "tri"}) {
			return 0, graphicsOutlineUnsupported("compound patch")
		}
	}
	if len(patch.Join) > 0 {
		target.Join = nil
		if e = graphicsDecodeOutline(patch.Join, &target.Join); e != nil {
			return 0, e
		}
		if e = graphicsOutlineJoin(target.Join); e != nil {
			return 0, e
		}
	}
	for _, spec := range []struct {
		raw json.RawMessage
		out **OutlineEnd
	}{{patch.HeadEnd, &target.HeadEnd}, {patch.TailEnd, &target.TailEnd}} {
		if len(spec.raw) == 0 {
			continue
		}
		*spec.out = nil
		if e = graphicsDecodeOutline(spec.raw, spec.out); e != nil {
			return 0, e
		}
		if e = graphicsOutlineEnd(*spec.out); e != nil {
			return 0, e
		}
	}
	if reflect.DeepEqual(old, target) {
		return 0, nil
	}
	next := source
	attrs := []losslessxml.AttributeEdit{}
	if target.Cap != old.Cap {
		attrs = append(attrs, losslessxml.AttributeEdit{Target: line, Name: name("", "cap"), Value: target.Cap})
	}
	if target.Compound != old.Compound {
		attrs = append(attrs, losslessxml.AttributeEdit{Target: line, Name: name("", "cmpd"), Value: target.Compound})
	}
	if len(attrs) > 0 {
		next, e = doc.Edit(nil, attrs)
		if e != nil {
			return 0, e
		}
	}
	for _, key := range []string{"join", "headEnd", "tailEnd"} {
		var equal, present bool
		var content string
		var join *OutlineJoin
		var end *OutlineEnd
		switch key {
		case "join":
			present = len(patch.Join) > 0
			equal = reflect.DeepEqual(old.Join, target.Join)
			join = target.Join
			if join != nil {
				limit := ""
				if join.Type == "miter" {
					limit = fmt.Sprintf(` lim="%d"`, *join.Limit)
				}
				content = fmt.Sprintf(`<a:%s xmlns:a="%s"%s/>`, join.Type, packaging.NSDrawingML, limit)
			}
		case "headEnd":
			present = len(patch.HeadEnd) > 0
			equal = reflect.DeepEqual(old.HeadEnd, target.HeadEnd)
			end = target.HeadEnd
		case "tailEnd":
			present = len(patch.TailEnd) > 0
			equal = reflect.DeepEqual(old.TailEnd, target.TailEnd)
			end = target.TailEnd
		}
		if !present || equal {
			continue
		}
		if end != nil {
			content = fmt.Sprintf(`<a:%s xmlns:a="%s" type="%s" w="%s" len="%s"/>`, key, packaging.NSDrawingML, end.Type, end.Width, end.Length)
		}
		d, ln, nodes, _, er := graphicsOutlineSelection(next, id)
		if er != nil {
			return 0, er
		}
		existing, exists := nodes[key]
		if exists && content != "" && (key != "join" || existing.Name().Local == join.Type) {
			changes := []losslessxml.AttributeEdit{}
			if key == "join" {
				if join.Type == "miter" {
					changes = append(changes, losslessxml.AttributeEdit{Target: existing, Name: name("", "lim"), Value: strconv.FormatInt(*join.Limit, 10)})
				}
			} else {
				for _, a := range []struct{ k, v string }{{"type", end.Type}, {"w", end.Width}, {"len", end.Length}} {
					changes = append(changes, losslessxml.AttributeEdit{Target: existing, Name: name("", a.k), Value: a.v})
				}
			}
			next, e = d.Edit(nil, changes)
		} else if exists {
			a, z := existing.SourceRange()
			next, e = fmtSplice(d, next, a, z, []byte(content))
		} else if content != "" {
			raw := ln.Raw()
			if strings.HasSuffix(string(raw), "/>") {
				a, z := ln.SourceRange()
				value := string(raw[:len(raw)-2]) + ">" + content + "</" + ln.QualifiedName() + ">"
				next, e = fmtSplice(d, next, a, z, []byte(value))
			} else {
				_, at := ln.ContentRange()
				following := []string{"extLst"}
				if key == "join" {
					following = []string{"headEnd", "tailEnd", "extLst"}
				} else if key == "headEnd" {
					following = []string{"tailEnd", "extLst"}
				}
				for _, n := range graphicsChildren(d, ln) {
					if graphicsChoice(n.Name().Local, following) {
						at, _ = n.SourceRange()
						break
					}
				}
				next, e = fmtSplice(d, next, at, at, []byte(content))
			}
		}
		if e != nil {
			return 0, e
		}
	}
	return s.graphicsMetadataCommit(part, h, source, next)
}
