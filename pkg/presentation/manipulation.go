package presentation

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ShapeTarget is an immutable, generation-bound snapshot of one ungrouped text shape.
// Mutating it requires its originating EditSession and a fresh part fingerprint.
type ShapeTarget struct {
	session     *EditSession
	generation  uint64
	part, hash  string
	id          uint32
	doc         *losslessxml.Document
	shape, body losslessxml.Element
	paragraphs  []losslessxml.Element
	text        string
	consumed    bool
}

func (t *ShapeTarget) Text() string {
	if t == nil {
		return ""
	}
	return t.text
}
func manipulationChildren(d *losslessxml.Document, parent losslessxml.Element, want xml.Name) []losslessxml.Element {
	var out []losslessxml.Element
	for _, e := range d.Elements() {
		if p, ok := e.Parent(); ok && p == parent && e.Name() == want {
			out = append(out, e)
		}
	}
	return out
}
func manipulationAttr(e losslessxml.Element, local string) string {
	for _, a := range e.Attributes() {
		if a.Name == (xml.Name{Local: local}) {
			return a.Value
		}
	}
	return ""
}
func manipulationWithin(e, ancestor losslessxml.Element) bool {
	for p, ok := e.Parent(); ok; p, ok = p.Parent() {
		if p == ancestor {
			return true
		}
	}
	return false
}
func (s *EditSession) manipulationProtection() error {
	g, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	for _, part := range g.Parts {
		if part.ContentType != packaging.ContentTypePresentation {
			continue
		}
		b, _, err := s.pkg.Part(part.Name)
		if err != nil {
			return err
		}
		d, err := losslessxml.Parse(b)
		if err != nil {
			return err
		}
		for _, e := range d.Elements() {
			if e.Name() == name(packaging.NSPresentationML, "modifyVerifier") {
				return editRefusal("protected_operation", "presentation modification protection")
			}
		}
	}
	return nil
}
func (s *EditSession) FindShape(part string, id uint32) (*ShapeTarget, error) {
	if !s.slides[part] {
		return nil, editRefusal("missing_target", "part is not an enrolled slide")
	}
	data, hash, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return nil, err
	}
	if len(doc.Elements()) == 0 || doc.Elements()[0].Name() != name(packaging.NSPresentationML, "sld") {
		return nil, editRefusal("unsupported_structure", "slide root required")
	}
	if err := s.manipulationProtection(); err != nil {
		return nil, err
	}
	var shape losslessxml.Element
	count := 0
	seenIDs := map[uint64]bool{}
	for _, e := range doc.Elements() {
		if e.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		n, er := strconv.ParseUint(manipulationAttr(e, "id"), 10, 32)
		if er != nil || n == 0 {
			return nil, editRefusal("unsupported_structure", "invalid shape ID")
		}
		if seenIDs[n] {
			return nil, editRefusal("ambiguous_target", "duplicate shape ID anywhere in slide")
		}
		seenIDs[n] = true
		if uint32(n) != id {
			continue
		}
		count++
		nv, ok := e.Parent()
		if !ok || nv.Name() != name(packaging.NSPresentationML, "nvSpPr") {
			return nil, editRefusal("unsupported_structure", "not an ordinary shape")
		}
		shape, ok = nv.Parent()
		if !ok || shape.Name() != name(packaging.NSPresentationML, "sp") {
			return nil, editRefusal("unsupported_structure", "not an ordinary shape")
		}
	}
	if count != 1 {
		kind := "missing_target"
		if count > 1 {
			kind = "ambiguous_target"
		}
		return nil, editRefusal(kind, "shape ID must resolve uniquely")
	}
	parent, ok := shape.Parent()
	if !ok || parent.Name() != name(packaging.NSPresentationML, "spTree") {
		return nil, editRefusal("unsupported_structure", "grouped shape")
	}
	common, ok := parent.Parent()
	if !ok || common.Name() != name(packaging.NSPresentationML, "cSld") {
		return nil, editRefusal("unsupported_structure", "shape tree has invalid owner")
	}
	root, ok := common.Parent()
	if !ok || root != doc.Elements()[0] {
		return nil, editRefusal("unsupported_structure", "common slide has invalid owner")
	}
	bodies := manipulationChildren(doc, shape, name(packaging.NSPresentationML, "txBody"))
	if len(bodies) != 1 {
		return nil, editRefusal("unsupported_structure", "one text body required")
	}
	body := bodies[0]
	if len(manipulationChildren(doc, body, name(packaging.NSDrawingML, "bodyPr"))) != 1 || len(manipulationChildren(doc, body, name(packaging.NSDrawingML, "lstStyle"))) != 1 {
		return nil, editRefusal("unsupported_structure", "body metadata required")
	}
	for _, e := range doc.Elements() {
		if e != shape && !manipulationWithin(e, shape) {
			continue
		}
		n := e.Name()
		if n.Local == "spLocks" {
			attrs := e.Attributes()
			if len(attrs) != 1 || attrs[0].Name != (xml.Name{Local: "noGrp"}) || attrs[0].Value != "1" {
				return nil, editRefusal("protected_operation", "unproved shape lock")
			}
			continue
		}
		if n.Local == "extLst" || n.Local == "fld" || n.Local == "br" || n.Local == "hlinkClick" || n.Local == "hlinkMouseOver" {
			return nil, editRefusal("unsupported_structure", "unsupported text structure")
		}
		if n.Space != packaging.NSPresentationML && n.Space != packaging.NSDrawingML {
			return nil, editRefusal("unsupported_structure", "foreign shape namespace")
		}
	}
	if err := manipulationTextLexicalBarrier(data, body); err != nil {
		return nil, err
	}
	paragraphs := manipulationChildren(doc, body, name(packaging.NSDrawingML, "p"))
	if len(paragraphs) == 0 {
		return nil, editRefusal("unsupported_structure", "at least one paragraph required")
	}
	seenParagraph := false
	seenList := false
	for _, e := range doc.Elements() {
		if parent, ok := e.Parent(); ok && parent == body {
			switch e.Name() {
			case name(packaging.NSDrawingML, "bodyPr"):
				if seenList || seenParagraph {
					return nil, editRefusal("unsupported_structure", "bodyPr out of order")
				}
			case name(packaging.NSDrawingML, "lstStyle"):
				if seenParagraph {
					return nil, editRefusal("unsupported_structure", "list style out of order")
				}
				seenList = true
			case name(packaging.NSDrawingML, "p"):
				if !seenList {
					return nil, editRefusal("unsupported_structure", "paragraph precedes list style")
				}
				seenParagraph = true
			default:
				return nil, editRefusal("unsupported_structure", "unexpected text body child")
			}
		}
		if !manipulationWithin(e, body) {
			continue
		}
		if parent, ok := e.Parent(); ok {
			switch e.Name() {
			case name(packaging.NSDrawingML, "pPr"), name(packaging.NSDrawingML, "r"), name(packaging.NSDrawingML, "endParaRPr"):
				if parent.Name() != name(packaging.NSDrawingML, "p") {
					return nil, editRefusal("unsupported_structure", "paragraph child owner invalid")
				}
			case name(packaging.NSDrawingML, "rPr"), name(packaging.NSDrawingML, "t"):
				if parent.Name() != name(packaging.NSDrawingML, "r") {
					return nil, editRefusal("unsupported_structure", "run child owner invalid")
				}
			case name(packaging.NSDrawingML, "buChar"), name(packaging.NSDrawingML, "buNone"), name(packaging.NSDrawingML, "buAutoNum"):
				if parent.Name() != name(packaging.NSDrawingML, "pPr") {
					return nil, editRefusal("unsupported_structure", "bullet owner invalid")
				}
			case name(packaging.NSDrawingML, "normAutofit"), name(packaging.NSDrawingML, "noAutofit"), name(packaging.NSDrawingML, "spAutoFit"):
				if parent.Name() != name(packaging.NSDrawingML, "bodyPr") {
					return nil, editRefusal("unsupported_structure", "autofit owner invalid")
				}
			}
		}
		switch e.Name() {
		case name(packaging.NSDrawingML, "p"), name(packaging.NSDrawingML, "pPr"), name(packaging.NSDrawingML, "r"), name(packaging.NSDrawingML, "rPr"), name(packaging.NSDrawingML, "t"), name(packaging.NSDrawingML, "bodyPr"), name(packaging.NSDrawingML, "lstStyle"), name(packaging.NSDrawingML, "buChar"), name(packaging.NSDrawingML, "buNone"), name(packaging.NSDrawingML, "buAutoNum"), name(packaging.NSDrawingML, "normAutofit"), name(packaging.NSDrawingML, "noAutofit"), name(packaging.NSDrawingML, "spAutoFit"), name(packaging.NSDrawingML, "endParaRPr"):
		default:
			return nil, editRefusal("unsupported_structure", "unsupported direct text topology")
		}
	}
	lines := make([]string, 0, len(paragraphs))
	for _, p := range paragraphs {
		pPrCount, endCount := 0, 0
		seenRun := false
		for _, child := range doc.Elements() {
			if parent, ok := child.Parent(); ok && parent == p {
				switch child.Name() {
				case name(packaging.NSDrawingML, "pPr"):
					pPrCount++
					if seenRun || endCount > 0 {
						return nil, editRefusal("unsupported_structure", "paragraph property after run")
					}
				case name(packaging.NSDrawingML, "endParaRPr"):
					endCount++
				case name(packaging.NSDrawingML, "r"):
					if endCount > 0 {
						return nil, editRefusal("unsupported_structure", "run after end properties")
					}
					seenRun = true
				}
			}
		}
		if pPrCount > 1 || endCount > 1 {
			return nil, editRefusal("unsupported_structure", "duplicate paragraph properties")
		}
		for _, run := range manipulationChildren(doc, p, name(packaging.NSDrawingML, "r")) {
			rPrCount, tCount := 0, 0
			seenText := false
			for _, child := range doc.Elements() {
				if parent, ok := child.Parent(); ok && parent == run {
					switch child.Name() {
					case name(packaging.NSDrawingML, "rPr"):
						rPrCount++
						if seenText {
							return nil, editRefusal("unsupported_structure", "run property after text")
						}
					case name(packaging.NSDrawingML, "t"):
						tCount++
						seenText = true
					}
				}
			}
			if rPrCount > 1 || tCount != 1 {
				return nil, editRefusal("unsupported_structure", "duplicate run property or missing text")
			}
		}
		var line strings.Builder
		for _, e := range doc.Elements() {
			if e.Name() == name(packaging.NSDrawingML, "t") && manipulationWithin(e, p) {
				text, leaf := e.Text()
				if !leaf {
					return nil, editRefusal("unsupported_structure", "mixed text leaf")
				}
				line.WriteString(text)
			}
		}
		lines = append(lines, line.String())
	}
	return &ShapeTarget{session: s, generation: s.generation, part: part, hash: hash, id: id, doc: doc, shape: shape, body: body, paragraphs: paragraphs, text: strings.Join(lines, "\n")}, nil
}
func manipulationTextLexicalBarrier(source []byte, body losslessxml.Element) error {
	start, end := body.ContentRange()
	if start < 0 || end > len(source) || start > end {
		return editRefusal("unsupported_structure", "invalid text-body range")
	}
	data := source[start:end]
	if bytes.Contains(data, []byte("<![CDATA[")) || bytes.Contains(data, []byte("<!--")) || bytes.Contains(data, []byte("<?")) || bytes.Contains(data, []byte("<!")) {
		return editRefusal("unsupported_structure", "lexical text-body barrier")
	}
	return nil
}
func (s *EditSession) shapeCurrent(t *ShapeTarget) error {
	if t == nil || t.session != s || t.consumed || t.generation != s.generation {
		return editRefusal("stale_target", "foreign, consumed or stale shape")
	}
	_, h, err := s.pkg.Part(t.part)
	if err != nil {
		return err
	}
	if h != t.hash {
		return editRefusal("stale_target", "shape part changed")
	}
	return nil
}
func (s *EditSession) shapeCommit(t *ShapeTarget, data []byte) error {
	if err := s.shapeCurrent(t); err != nil {
		return err
	}
	before, _, err := s.pkg.Part(t.part)
	if err != nil {
		return err
	}
	if bytes.Equal(before, data) {
		return nil
	}
	if err := s.pkg.Replace([]packaging.Replacement{{Part: t.part, ExpectedSHA256: t.hash, Data: data}}); err != nil {
		return err
	}
	s.generation++
	t.consumed = true
	return nil
}
func manipulationParagraph(text string) losslessxml.NewElement {
	a := packaging.NSDrawingML
	return losslessxml.NewElement{Name: name(a, "p"), Children: []losslessxml.NewElement{{Name: name(a, "r"), Children: []losslessxml.NewElement{{Name: name(a, "t"), Text: text}}}}}
}
func manipulationXMLText(text string) error {
	if strings.ContainsAny(text, "\r\t") {
		return editRefusal("unsupported_structure", "CR and tabs require explicit text semantics")
	}
	return nil
}
func manipulationReplaceParagraphs(t *ShapeTarget, paras []losslessxml.NewElement) ([]byte, error) {
	edits := []losslessxml.ElementReplacement{{Target: t.paragraphs[0], Nodes: paras}}
	for _, p := range t.paragraphs[1:] {
		edits = append(edits, losslessxml.ElementReplacement{Target: p})
	}
	return t.doc.ReplaceElements(edits)
}

// SetShapeText replaces the whole plain text frame, or appends LF-separated
// paragraphs. Existing paragraph formatting is deliberately reset on replacement.
func (s *EditSession) SetShapeText(t *ShapeTarget, text string, appendText bool) error {
	if err := s.shapeCurrent(t); err != nil {
		return err
	}
	if err := manipulationXMLText(text); err != nil {
		return err
	}
	lines := strings.Split(text, "\n")
	paras := make([]losslessxml.NewElement, 0, len(lines))
	for _, line := range lines {
		paras = append(paras, manipulationParagraph(line))
	}
	var data []byte
	var err error
	if appendText {
		data, err = t.doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: t.body, Children: paras}})
	} else {
		data, err = manipulationReplaceParagraphs(t, paras)
	}
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	return s.shapeCommit(t, data)
}

// ClearShapeText leaves one empty paragraph without bullet properties.
func (s *EditSession) ClearShapeText(t *ShapeTarget) error {
	if err := s.shapeCurrent(t); err != nil {
		return err
	}
	data, err := manipulationReplaceParagraphs(t, []losslessxml.NewElement{{Name: name(packaging.NSDrawingML, "p")}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	return s.shapeCommit(t, data)
}

// AddShapeBullet appends an explicit bullet. A label is the bold first run;
// the description is a separate unbolded run.
func (s *EditSession) AddShapeBullet(t *ShapeTarget, text string, level int, label string) error {
	if err := s.shapeCurrent(t); err != nil {
		return err
	}
	if level < 0 || level > 8 {
		return editRefusal("unsupported_structure", "bullet level outside 0..8")
	}
	if strings.ContainsAny(text, "\r\n") || strings.ContainsAny(label, "\r\n") {
		return editRefusal("unsupported_structure", "one paragraph per bullet")
	}
	if err := manipulationXMLText(text); err != nil {
		return err
	}
	if err := manipulationXMLText(label); err != nil {
		return err
	}
	a := packaging.NSDrawingML
	p := losslessxml.NewElement{Name: name(a, "p"), Children: []losslessxml.NewElement{{Name: name(a, "pPr"), Attributes: []xml.Attr{{Name: xml.Name{Local: "lvl"}, Value: strconv.Itoa(level)}}, Children: []losslessxml.NewElement{{Name: name(a, "buChar"), Attributes: []xml.Attr{{Name: xml.Name{Local: "char"}, Value: "•"}}}}}}}
	if label != "" {
		p.Children = append(p.Children, losslessxml.NewElement{Name: name(a, "r"), Children: []losslessxml.NewElement{{Name: name(a, "rPr"), Attributes: []xml.Attr{{Name: xml.Name{Local: "b"}, Value: "1"}}}, {Name: name(a, "t"), Text: label + ": "}}})
	}
	p.Children = append(p.Children, losslessxml.NewElement{Name: name(a, "r"), Children: []losslessxml.NewElement{{Name: name(a, "t"), Text: text}}})
	var data []byte
	var err error
	if len(t.paragraphs) == 1 && t.text == "" {
		data, err = manipulationReplaceParagraphs(t, []losslessxml.NewElement{p})
	} else {
		data, err = t.doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: t.body, Children: []losslessxml.NewElement{p}}})
	}
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	return s.shapeCommit(t, data)
}

// SetShapeAutofit replaces only the bodyPr fit choice. Other body attributes
// and the text frame's list style and paragraphs are retained.
func (s *EditSession) SetShapeAutofit(t *ShapeTarget, mode string) error {
	if err := s.shapeCurrent(t); err != nil {
		return err
	}
	child := map[string]string{"shrink": "normAutofit", "none": "noAutofit", "resize": "spAutoFit"}[mode]
	if child == "" {
		return editRefusal("unsupported_structure", "unknown autofit mode")
	}
	bodyPr := manipulationChildren(t.doc, t.body, name(packaging.NSDrawingML, "bodyPr"))[0]
	fitCount := 0
	for _, e := range t.doc.Elements() {
		if p, ok := e.Parent(); ok && p == bodyPr {
			if e.Name() != name(packaging.NSDrawingML, "normAutofit") && e.Name() != name(packaging.NSDrawingML, "noAutofit") && e.Name() != name(packaging.NSDrawingML, "spAutoFit") {
				return editRefusal("unsupported_structure", "unknown bodyPr child")
			}
			fitCount++
		}
	}
	if fitCount > 1 {
		return editRefusal("unsupported_structure", "ambiguous autofit choices")
	}
	data, err := t.doc.ReplaceElements([]losslessxml.ElementReplacement{{Target: bodyPr, Nodes: []losslessxml.NewElement{{Name: bodyPr.Name(), Attributes: bodyPr.Attributes(), Children: []losslessxml.NewElement{{Name: name(packaging.NSDrawingML, child)}}}}}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	return s.shapeCommit(t, data)
}

// TableValues authors a new plain DrawingML table in one enrolled slide. All
// dimensions and cell values are preflighted before replacing the slide part.
func (s *EditSession) TableValues(part string, rows [][]string, x, y, width, height int64) error {
	if !s.slides[part] {
		return editRefusal("missing_target", "slide not enrolled")
	}
	if err := s.manipulationProtection(); err != nil {
		return err
	}
	if len(rows) == 0 || len(rows) > 64 || len(rows[0]) == 0 || len(rows[0]) > 64 || width <= 0 || height <= 0 || x < 0 || y < 0 || width < int64(len(rows[0])) || height < int64(len(rows)) || width > 100000000 || height > 100000000 || x > 100000000-width || y > 100000000-height {
		return editRefusal("unsupported_structure", "table geometry or dimensions invalid")
	}
	cols := len(rows[0])
	cellBytes := 0
	for _, row := range rows {
		if len(row) != cols {
			return editRefusal("unsupported_structure", "ragged table")
		}
		for _, v := range row {
			if err := manipulationXMLText(v); err != nil {
				return err
			}
			if len(v) > 4096 || cellBytes > 65536-len(v) {
				return editRefusal("unsupported_structure", "table cell text budget exceeded")
			}
			cellBytes += len(v)
		}
	}
	b, h, err := s.pkg.Part(part)
	if err != nil {
		return err
	}
	d, err := losslessxml.Parse(b)
	if err != nil {
		return err
	}
	if len(d.Elements()) == 0 || d.Elements()[0].Name() != name(packaging.NSPresentationML, "sld") {
		return editRefusal("unsupported_structure", "slide root required")
	}
	var tree losslessxml.Element
	count := 0
	maxID := uint64(0)
	seenIDs := map[uint64]bool{}
	for _, e := range d.Elements() {
		if e.Name() == name(packaging.NSPresentationML, "spTree") {
			parent, ok := e.Parent()
			root, hasRoot := parent.Parent()
			if !ok || parent.Name() != name(packaging.NSPresentationML, "cSld") || !hasRoot || root != d.Elements()[0] {
				return editRefusal("unsupported_structure", "shape tree owner invalid")
			}
			tree = e
			count++
		}
		if e.Name() == name(packaging.NSPresentationML, "cNvPr") {
			id, er := strconv.ParseUint(manipulationAttr(e, "id"), 10, 32)
			if er != nil || id == 0 {
				return editRefusal("unsupported_structure", "invalid shape ID")
			}
			if seenIDs[id] {
				return editRefusal("ambiguous_target", "duplicate shape ID")
			}
			seenIDs[id] = true
			if id > maxID {
				maxID = id
			}
		}
	}
	if count != 1 || maxID >= 1<<32-1 {
		return editRefusal("unsupported_structure", "ambiguous shape tree or ID exhausted")
	}
	if err := manipulationTextLexicalBarrier(b, tree); err != nil {
		return err
	}
	p, a := packaging.NSPresentationML, packaging.NSDrawingML
	elem := func(ns, local string, attrs []xml.Attr, children ...losslessxml.NewElement) losslessxml.NewElement {
		return losslessxml.NewElement{Name: name(ns, local), Attributes: attrs, Children: children}
	}
	attr := func(k, v string) xml.Attr { return xml.Attr{Name: xml.Name{Local: k}, Value: v} }
	grid := []losslessxml.NewElement{}
	for c := 0; c < cols; c++ {
		w := width / int64(cols)
		if c == cols-1 {
			w += width % int64(cols)
		}
		grid = append(grid, elem(a, "gridCol", []xml.Attr{attr("w", strconv.FormatInt(w, 10))}))
	}
	tbl := []losslessxml.NewElement{elem(a, "tblPr", nil), elem(a, "tblGrid", nil, grid...)}
	for r, row := range rows {
		cells := []losslessxml.NewElement{}
		for _, v := range row {
			cells = append(cells, elem(a, "tc", nil, elem(a, "txBody", nil, elem(a, "bodyPr", nil), elem(a, "lstStyle", nil), manipulationParagraph(v)), elem(a, "tcPr", nil)))
		}
		rh := height / int64(len(rows))
		if r == len(rows)-1 {
			rh += height % int64(len(rows))
		}
		tbl = append(tbl, elem(a, "tr", []xml.Attr{attr("h", strconv.FormatInt(rh, 10))}, cells...))
	}
	id := strconv.FormatUint(maxID+1, 10)
	frame := elem(p, "graphicFrame", nil, elem(p, "nvGraphicFramePr", nil, elem(p, "cNvPr", []xml.Attr{attr("id", id), attr("name", "Table "+id)}), elem(p, "cNvGraphicFramePr", nil), elem(p, "nvPr", nil)), elem(p, "xfrm", nil, elem(a, "off", []xml.Attr{attr("x", strconv.FormatInt(x, 10)), attr("y", strconv.FormatInt(y, 10))}), elem(a, "ext", []xml.Attr{attr("cx", strconv.FormatInt(width, 10)), attr("cy", strconv.FormatInt(height, 10))})), elem(a, "graphic", nil, elem(a, "graphicData", []xml.Attr{attr("uri", "http://schemas.openxmlformats.org/drawingml/2006/table")}, elem(a, "tbl", nil, tbl...))))
	data, err := d.InsertChildren([]losslessxml.ChildInsertion{{Parent: tree, Children: []losslessxml.NewElement{frame}}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: h, Data: data}}); err != nil {
		return err
	}
	s.generation++
	return nil
}

// ReorderSlides changes only the sldId list, retaining every slide part and
// relationship. A permutation is zero-based and is fully checked first.
func (s *EditSession) ReorderSlides(order []int) error {
	ids, main, doc, hash, err := s.slideIDs()
	if err != nil {
		return err
	}
	n := len(ids)
	if len(order) != n {
		return editRefusal("invalid-permutation", "wrong number of slides")
	}
	seen := make([]bool, n)
	for _, v := range order {
		if v < 0 || v >= n || seen[v] {
			return editRefusal("invalid-permutation", "not a permutation")
		}
		seen[v] = true
	}
	list, _ := ids[0].Parent()
	children := make([]losslessxml.NewElement, 0, n)
	for _, idx := range order {
		children = append(children, losslessxml.NewElement{Name: ids[idx].Name(), Attributes: ids[idx].Attributes()})
	}
	b, err := doc.ReplaceElements([]losslessxml.ElementReplacement{{Target: list, Nodes: []losslessxml.NewElement{{Name: list.Name(), Attributes: list.Attributes(), Children: children}}}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	current, _, err := s.pkg.Part(main)
	if err != nil {
		return err
	}
	if bytes.Equal(b, current) {
		return nil
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: main, ExpectedSHA256: hash, Data: b}}); err != nil {
		return err
	}
	s.generation++
	return nil
}
func (s *EditSession) slideIDs() ([]losslessxml.Element, string, *losslessxml.Document, string, error) {
	if err := s.manipulationProtection(); err != nil {
		return nil, "", nil, "", err
	}
	g, err := s.pkg.Graph()
	if err != nil {
		return nil, "", nil, "", err
	}
	main := ""
	for _, edge := range g.Edges {
		if edge.Source == "" && edge.Type == packaging.RelTypeOfficeDocument && !edge.External {
			if main != "" {
				return nil, "", nil, "", editRefusal("ambiguous_target", "multiple main parts")
			}
			main = edge.ResolvedPart
		}
	}
	if main == "" {
		return nil, "", nil, "", editRefusal("missing_target", "presentation main part")
	}
	b, h, err := s.pkg.Part(main)
	if err != nil {
		return nil, "", nil, "", err
	}
	d, err := losslessxml.Parse(b)
	if err != nil {
		return nil, "", nil, "", err
	}
	if len(d.Elements()) == 0 || d.Elements()[0].Name() != name(packaging.NSPresentationML, "presentation") {
		return nil, "", nil, "", editRefusal("unsupported_structure", "presentation root required")
	}
	var list losslessxml.Element
	count := 0
	for _, e := range d.Elements() {
		if e.Name() == name(packaging.NSPresentationML, "sldIdLst") {
			parent, ok := e.Parent()
			if !ok || parent != d.Elements()[0] {
				return nil, "", nil, "", editRefusal("unsupported_structure", "nested slide list")
			}
			list = e
			count++
		}
		if e.Name() == name(packaging.NSPresentationML, "custShowLst") || e.Name() == name(packaging.NSPresentationML, "extLst") || e.Name() == name(packaging.NSPresentationML, "sectionLst") {
			return nil, "", nil, "", editRefusal("unsupported_structure", "slide order metadata requires explicit policy")
		}
	}
	if count != 1 {
		return nil, "", nil, "", editRefusal("unsupported_structure", "one slide list required")
	}
	if len(list.Attributes()) != 0 {
		return nil, "", nil, "", editRefusal("unsupported_structure", "slide list has unproved attributes")
	}
	if err := manipulationTextLexicalBarrier(b, list); err != nil {
		return nil, "", nil, "", err
	}
	for _, e := range d.Elements() {
		if parent, ok := e.Parent(); ok && parent == list && e.Name() != name(packaging.NSPresentationML, "sldId") {
			return nil, "", nil, "", editRefusal("unsupported_structure", "unexpected slide-list child")
		}
	}
	ids := manipulationChildren(d, list, name(packaging.NSPresentationML, "sldId"))
	listStart, listEnd := list.ContentRange()
	cursor := listStart
	for _, entry := range ids {
		start, end := entry.SourceRange()
		if start != cursor || end < start || end > listEnd {
			return nil, "", nil, "", editRefusal("unsupported_structure", "slide list lexical gap")
		}
		cursor = end
	}
	if cursor != listEnd {
		return nil, "", nil, "", editRefusal("unsupported_structure", "slide list trailing lexical gap")
	}
	if len(ids) == 0 || len(ids) != len(s.slides) {
		return nil, "", nil, "", editRefusal("unsupported_structure", "unresolved slide list")
	}
	rels := map[string]string{}
	for _, edge := range g.Edges {
		if edge.Source == main && edge.Type == packaging.RelTypeSlide && !edge.External {
			rels[edge.ID] = edge.ResolvedPart
		}
	}
	used := map[string]bool{}
	usedNumeric := map[uint64]bool{}
	usedParts := map[string]bool{}
	for _, id := range ids {
		if len(id.Attributes()) != 2 {
			return nil, "", nil, "", editRefusal("unsupported_structure", "unexpected slide ID attributes")
		}
		seenAttr := map[xml.Name]bool{}
		for _, attr := range id.Attributes() {
			if attr.Name != (xml.Name{Local: "id"}) && attr.Name != name(packaging.NSDocumentRelationships, "id") {
				return nil, "", nil, "", editRefusal("unsupported_structure", "unsupported slide ID attribute")
			}
			seenAttr[attr.Name] = true
		}
		if len(seenAttr) != 2 {
			return nil, "", nil, "", editRefusal("unsupported_structure", "slide ID attributes missing")
		}
		for _, e := range d.Elements() {
			if parent, ok := e.Parent(); ok && parent == id {
				return nil, "", nil, "", editRefusal("unsupported_structure", "nested slide ID")
			}
		}
		numeric, err := strconv.ParseUint(manipulationAttr(id, "id"), 10, 32)
		if err != nil || numeric < 256 || numeric > 2147483647 || usedNumeric[numeric] {
			return nil, "", nil, "", editRefusal("ambiguous_target", "invalid or duplicate slide ID")
		}
		usedNumeric[numeric] = true
		rid := ""
		for _, a := range id.Attributes() {
			if a.Name == name(packaging.NSDocumentRelationships, "id") {
				rid = a.Value
			}
		}
		part := rels[rid]
		if rid == "" || part == "" || used[rid] || usedParts[part] || !s.slides[part] {
			return nil, "", nil, "", editRefusal("unsupported_structure", "ambiguous slide edge or part")
		}
		used[rid] = true
		usedParts[part] = true
	}
	if len(used) != len(rels) {
		return nil, "", nil, "", editRefusal("unsupported_structure", "orphan slide relationship")
	}
	return ids, main, d, h, nil
}

// InsertTextSlide inserts an owned title/subtitle slide using one complete OPC
// graph plan. It clones only the safe slide payload of the final plain slide;
// no existing slide payload or relationship is renamed or rewritten.
func (s *EditSession) InsertTextSlide(index int, title, subtitle string) error {
	if err := manipulationXMLText(title); err != nil {
		return err
	}
	if err := manipulationXMLText(subtitle); err != nil {
		return err
	}
	ids, main, d, h, err := s.slideIDs()
	if err != nil {
		return err
	}
	if index < 0 || index > len(ids) {
		return editRefusal("invalid-permutation", "insertion index out of range")
	}
	g, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	rels := map[string]string{}
	allRID := map[string]bool{}
	for _, e := range g.Edges {
		if e.Source == main {
			allRID[e.ID] = true
			if e.Type == packaging.RelTypeSlide {
				rels[e.ID] = e.ResolvedPart
			}
		}
	}
	lastRID := ""
	for _, a := range ids[len(ids)-1].Attributes() {
		if a.Name == name(packaging.NSDocumentRelationships, "id") {
			lastRID = a.Value
		}
	}
	template := rels[lastRID]
	if template == "" {
		return editRefusal("unsupported_structure", "no title slide template")
	}
	tb, _, err := s.pkg.Part(template)
	if err != nil {
		return err
	}
	td, err := losslessxml.Parse(tb)
	if err != nil {
		return err
	}
	// The template must contain only title/subtitle shapes and an existing layout
	// edge. Otherwise cloning could duplicate an unowned relationship or shape.
	if len(td.Elements()) == 0 || td.Elements()[0].Name() != name(packaging.NSPresentationML, "sld") {
		return editRefusal("unsupported_structure", "template slide root")
	}
	if len(manipulationChildren(td, td.Elements()[0], name(packaging.NSPresentationML, "cSld"))) != 1 {
		return editRefusal("unsupported_structure", "template has ambiguous common slide")
	}
	// Choose the two direct text placeholders by expanded p:ph type and
	// distinct shape IDs. Text spelling is never an identity predicate.
	placeholderIDs := map[string]uint32{}
	seenShapes := map[uint64]bool{}
	shapeCount := 0
	for _, e := range td.Elements() {
		if e.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		numeric, er := strconv.ParseUint(manipulationAttr(e, "id"), 10, 32)
		if er != nil || numeric == 0 {
			return editRefusal("unsupported_structure", "invalid template shape ID")
		}
		if seenShapes[numeric] {
			return editRefusal("ambiguous_target", "duplicate template shape ID")
		}
		seenShapes[numeric] = true
		parent, ok := e.Parent()
		if !ok {
			return editRefusal("unsupported_structure", "template shape owner")
		}
		if parent.Name() != name(packaging.NSPresentationML, "nvSpPr") {
			if numeric != 1 || parent.Name() != name(packaging.NSPresentationML, "nvGrpSpPr") {
				return editRefusal("unsupported_structure", "template has non-text object")
			}
			continue
		}
		shape, ok := parent.Parent()
		if !ok || shape.Name() != name(packaging.NSPresentationML, "sp") {
			return editRefusal("unsupported_structure", "template shape topology")
		}
		owner, ok := shape.Parent()
		if !ok || owner.Name() != name(packaging.NSPresentationML, "spTree") {
			return editRefusal("unsupported_structure", "grouped template shape")
		}
		shapeCount++
		var ph losslessxml.Element
		for _, node := range td.Elements() {
			if node.Name() == name(packaging.NSPresentationML, "ph") && manipulationWithin(node, parent) {
				if ph.Ordinal() >= 0 {
					return editRefusal("ambiguous_target", "multiple placeholder types")
				}
				ph = node
			}
		}
		if ph.Ordinal() < 0 {
			return editRefusal("unsupported_structure", "template shape lacks placeholder")
		}
		typ := manipulationAttr(ph, "type")
		if typ == "ctrTitle" || typ == "title" {
			typ = "title"
		}
		if typ != "title" && typ != "subTitle" {
			return editRefusal("unsupported_structure", "template has non-title placeholder")
		}
		if _, exists := placeholderIDs[typ]; exists {
			return editRefusal("ambiguous_target", "duplicate template placeholder")
		}
		placeholderIDs[typ] = uint32(numeric)
	}
	if shapeCount != 2 || len(placeholderIDs) != 2 || len(seenShapes) != 3 {
		return editRefusal("unsupported_structure", "template needs exactly title/subtitle shapes")
	}
	for _, target := range []uint32{placeholderIDs["title"], placeholderIDs["subTitle"]} {
		if _, err := s.FindShape(template, target); err != nil {
			return err
		}
	}
	layout := ""
	for _, e := range g.Edges {
		if e.Source == template {
			if e.Type != packaging.RelTypeSlideLayout || e.External || layout != "" || e.ID != "rId1" {
				return editRefusal("relationship_policy", "template slide has unsupported relationships")
			}
			layout = e.ResolvedPart
		}
	}
	if layout == "" {
		return editRefusal("missing_target", "template layout")
	}
	newPart := ""
	for i := 1; i < 10000; i++ {
		candidate := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		found := false
		for _, part := range g.Parts {
			if strings.EqualFold(part.Name, candidate) {
				found = true
				break
			}
		}
		if !found {
			newPart = candidate
			break
		}
	}
	if newPart == "" {
		return editRefusal("unsupported_structure", "slide part namespace exhausted")
	}
	newRID := ""
	for i := 1; i < 10000; i++ {
		candidate := fmt.Sprintf("rId%d", i)
		if !allRID[candidate] {
			newRID = candidate
			break
		}
	}
	if newRID == "" {
		return editRefusal("unsupported_structure", "slide relationship IDs exhausted")
	}
	maxID := uint64(0)
	for _, e := range ids {
		id, er := strconv.ParseUint(manipulationAttr(e, "id"), 10, 32)
		if er != nil {
			return editRefusal("unsupported_structure", "invalid slide ID")
		}
		if id > maxID {
			maxID = id
		}
	}
	if maxID >= 2147483647 {
		return editRefusal("unsupported_structure", "slide ID exhausted")
	}
	// Author the inserted slide from the two placeholder owners. Replace the
	// complete paragraph list, including a formerly empty/self-closing title.
	titleShape, subtitleShape := losslessxml.Element{}, losslessxml.Element{}
	for _, e := range td.Elements() {
		if e.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		numeric, _ := strconv.ParseUint(manipulationAttr(e, "id"), 10, 32)
		if uint32(numeric) != placeholderIDs["title"] && uint32(numeric) != placeholderIDs["subTitle"] {
			continue
		}
		p, ok := e.Parent()
		if !ok {
			return editRefusal("unsupported_structure", "template owner disappeared")
		}
		shape, ok := p.Parent()
		if !ok {
			return editRefusal("unsupported_structure", "template owner disappeared")
		}
		if uint32(numeric) == placeholderIDs["title"] {
			titleShape = shape
		} else {
			subtitleShape = shape
		}
	}
	updates := []losslessxml.ElementReplacement{}
	for _, item := range []struct {
		shape losslessxml.Element
		text  string
	}{{titleShape, title}, {subtitleShape, subtitle}} {
		bodies := manipulationChildren(td, item.shape, name(packaging.NSPresentationML, "txBody"))
		if len(bodies) != 1 {
			return editRefusal("unsupported_structure", "template text body missing")
		}
		paragraphs := manipulationChildren(td, bodies[0], name(packaging.NSDrawingML, "p"))
		if len(paragraphs) == 0 {
			return editRefusal("unsupported_structure", "template paragraph missing")
		}
		updates = append(updates, losslessxml.ElementReplacement{Target: paragraphs[0], Nodes: []losslessxml.NewElement{manipulationParagraph(item.text)}})
		for _, para := range paragraphs[1:] {
			updates = append(updates, losslessxml.ElementReplacement{Target: para})
		}
	}
	newData, err := td.ReplaceElements(updates)
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	list, _ := ids[0].Parent()
	children := make([]losslessxml.NewElement, 0, len(ids)+1)
	for i := 0; i <= len(ids); i++ {
		if i == index {
			children = append(children, losslessxml.NewElement{Name: name(packaging.NSPresentationML, "sldId"), Attributes: []xml.Attr{{Name: xml.Name{Local: "id"}, Value: strconv.FormatUint(maxID+1, 10)}, {Name: name(packaging.NSDocumentRelationships, "id"), Value: newRID}}})
		}
		if i < len(ids) {
			children = append(children, losslessxml.NewElement{Name: ids[i].Name(), Attributes: ids[i].Attributes()})
		}
	}
	updated, err := d.ReplaceElements([]losslessxml.ElementReplacement{{Target: list, Nodes: []losslessxml.NewElement{{Name: list.Name(), Attributes: list.Attributes(), Children: children}}}})
	if err != nil {
		return editRefusal("unsupported_structure", err.Error())
	}
	change := packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: newPart, ContentType: packaging.ContentTypeSlide, Data: newData}}, Relationships: []packaging.RelationshipAddition{{Source: main, ID: newRID, Type: packaging.RelTypeSlide, TargetPart: newPart}, {Source: newPart, ID: "rId1", Type: packaging.RelTypeSlideLayout, TargetPart: layout}}, Replacements: []packaging.Replacement{{Part: main, ExpectedSHA256: h, Data: updated}}}
	plan, err := s.pkg.PlanGraphMutation(change)
	if err != nil {
		return err
	}
	if err := s.pkg.ApplyGraphPlan(plan); err != nil {
		return err
	}
	s.slides[newPart] = true
	s.generation++
	return nil
}
