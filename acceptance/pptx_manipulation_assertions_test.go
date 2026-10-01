package acceptance

import (
	"bytes"
	"fmt"
	"os"
	"reflect"
	"strconv"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Prove the custody checker rejects drift outside the selected text or
// added table frame, rather than merely asserting a successful edit.
func TestPPTXManipulationCustodyOracleFault(t *testing.T) {
	fixture, err := os.ReadFile(testutil.ReferencePath(pptxManipulationFixtureFile))
	if err != nil {
		t.Skipf("candidate fixture unavailable: %v", err)
	}
	members, err := pptxMembers(fixture)
	if err != nil {
		t.Fatal(err)
	}
	original := members["ppt/slides/slide1.xml"]
	if len(original) == 0 {
		t.Fatal("source slide absent")
	}
	changed := bytes.Replace(original, []byte(">Alpha</a:t>"), []byte(">Renamed</a:t>"), 1)
	if bytes.Equal(changed, original) {
		t.Fatal("selected title not found")
	}
	if err := pptxMaskedXMLSame(original, changed, "patch-title", "ppt/slides/slide1.xml", 2); err != nil {
		t.Fatal(err)
	}
	fault := bytes.Replace(changed, []byte("name=\"Title 1\""), []byte("name=\"Wrong 1\""), 1)
	if bytes.Equal(fault, changed) {
		t.Fatal("nonselected fault not injected")
	}
	if err := pptxMaskedXMLSame(original, fault, "patch-title", "ppt/slides/slide1.xml", 2); err == nil {
		t.Fatal("text oracle accepted nonselected XML drift")
	}
	frame := []byte(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Table 9"/></p:nvGraphicFramePr></p:graphicFrame>`)
	table := bytes.Replace(original, []byte("</p:spTree>"), append(frame, []byte("</p:spTree>")...), 1)
	if bytes.Equal(table, original) {
		t.Fatal("table injection absent")
	}
	if err := pptxAddedChildSame(original, table, "graphicFrame", "spTree", "", ""); err != nil {
		t.Fatal(err)
	}
	tableFault := bytes.Replace(table, []byte("name=\"Title 1\""), []byte("name=\"Wrong 1\""), 1)
	if err := pptxAddedChildSame(original, tableFault, "graphicFrame", "spTree", "", ""); err == nil {
		t.Fatal("table oracle accepted original shape drift")
	}
}

func pptxInt(x any) (int64, error) {
	v, ok := x.(float64)
	if !ok || float64(int64(v)) != v {
		return 0, fmt.Errorf("invalid integer %v", x)
	}
	return int64(v), nil
}
func pptxExpectedStrings(expected map[string]any, key string, got []string) error {
	want, exists := expected[key]
	if !exists {
		return nil
	}
	return pptxSameStrings(got, want)
}
func pptxSlideTitles(archive []byte) ([]string, []string, []string, []string, error) {
	paths, rids, ids, err := pptxSlideParts(archive)
	if err != nil {
		return nil, nil, nil, nil, err
	}
	p, err := pptxMembers(archive)
	if err != nil {
		return nil, nil, nil, nil, err
	}
	titles := []string{}
	for _, part := range paths {
		d, err := losslessxml.Parse(p[part])
		if err != nil {
			return nil, nil, nil, nil, err
		}
		sp, err := pptxShape(d, 2)
		if err != nil {
			return nil, nil, nil, nil, err
		}
		_, lines, err := pptxShapeParagraphs(d, sp)
		if err != nil {
			return nil, nil, nil, nil, err
		}
		titles = append(titles, strings.Join(lines, "\n"))
	}
	return paths, rids, ids, titles, nil
}
func pptxEdgeMap(archive []byte) (map[string]packaging.Edge, error) {
	p, err := packaging.OpenPreserved(archive, packaging.Limits{})
	if err != nil {
		return nil, err
	}
	g, err := p.Graph()
	if err != nil {
		return nil, err
	}
	out := map[string]packaging.Edge{}
	for _, e := range g.Edges {
		key := e.Source + "\x00" + e.ID
		if _, ok := out[key]; ok {
			return nil, fmt.Errorf("duplicate graph edge %s", key)
		}
		out[key] = e
	}
	return out, nil
}

// Only the declared byte range may change. Everything else, including root
// attributes, namespace declarations, original siblings, and PI/comment bytes,
// is compared directly against the independent source archive.
func pptxMaskedXMLSame(before, after []byte, kind, member string, shapeID uint32) error {
	source, err := losslessxml.Parse(before)
	if err != nil {
		return err
	}
	saved, err := losslessxml.Parse(after)
	if err != nil {
		return err
	}
	selectRange := func(doc *losslessxml.Document, blob []byte) (int, int, error) {
		var target losslessxml.Element
		switch {
		case member == "ppt/presentation.xml":
			for _, e := range doc.Elements() {
				if e.Name() == pptxElem(pptxP, "sldIdLst") {
					target = e
					break
				}
			}
			if target.Ordinal() < 0 {
				return 0, 0, fmt.Errorf("slide list absent")
			}
			a, b := target.ContentRange()
			return a, b, nil
		case member == "ppt/notesSlides/notesSlide1.xml":
			for _, e := range doc.Elements() {
				if e.Name() != pptxElem(pptxP, "ph") || pptxAttr(e, "", "type") != "body" {
					continue
				}
				for p, ok := e.Parent(); ok; p, ok = p.Parent() {
					if p.Name() == pptxElem(pptxP, "sp") {
						bodies := pptxChild(doc, p, pptxElem(pptxP, "txBody"))
						if len(bodies) != 1 {
							return 0, 0, fmt.Errorf("notes body count")
						}
						target = bodies[0]
						break
					}
				}
			}
			if target.Ordinal() < 0 {
				return 0, 0, fmt.Errorf("notes body absent")
			}
			ps := pptxChild(doc, target, pptxElem(pptxA, "p"))
			if len(ps) == 0 {
				return 0, 0, fmt.Errorf("notes paragraph absent")
			}
			a, _ := ps[0].SourceRange()
			_, b := ps[len(ps)-1].SourceRange()
			return a, b, nil
		case strings.HasPrefix(kind, "autofit-"):
			sp, e := pptxShape(doc, shapeID)
			if e != nil {
				return 0, 0, e
			}
			body := pptxChild(doc, sp, pptxElem(pptxP, "txBody"))
			if len(body) != 1 {
				return 0, 0, fmt.Errorf("text body absent")
			}
			props := pptxChild(doc, body[0], pptxElem(pptxA, "bodyPr"))
			if len(props) != 1 {
				return 0, 0, fmt.Errorf("bodyPr absent")
			}
			a, b := props[0].SourceRange()
			return a, b, nil
		default:
			sp, e := pptxShape(doc, shapeID)
			if e != nil {
				return 0, 0, e
			}
			body := pptxChild(doc, sp, pptxElem(pptxP, "txBody"))
			if len(body) != 1 {
				return 0, 0, fmt.Errorf("text body absent")
			}
			ps := pptxChild(doc, body[0], pptxElem(pptxA, "p"))
			if len(ps) == 0 {
				return 0, 0, fmt.Errorf("paragraph absent")
			}
			a, _ := ps[0].SourceRange()
			_, b := ps[len(ps)-1].SourceRange()
			return a, b, nil
		}
	}
	a, b, err := selectRange(source, before)
	if err != nil {
		return err
	}
	c, d, err := selectRange(saved, after)
	if err != nil {
		return err
	}
	if a < 0 || b < a || b > len(before) || c < 0 || d < c || d > len(after) {
		return fmt.Errorf("invalid selected source ranges")
	}
	if !bytes.Equal(before[:a], after[:c]) || !bytes.Equal(before[b:], after[d:]) {
		return fmt.Errorf("nonselected XML bytes changed in %s", member)
	}
	return nil
}

// XML registries only permit insertion of one selected child. Strip that new
// child from the saved bytes and demand the original registry bytes exactly.
func pptxAddedSlidePart(before, after map[string][]byte) string {
	found := ""
	for part := range after {
		if _, old := before[part]; !old && strings.HasPrefix(part, "ppt/slides/slide") && strings.HasSuffix(part, ".xml") {
			if found != "" {
				return ""
			}
			found = part
		}
	}
	return found
}
func pptxAddedChildSame(before, after []byte, element, parent, attribute, value string) error {
	doc, err := losslessxml.Parse(after)
	if err != nil {
		return err
	}
	matches := []losslessxml.Element{}
	for _, e := range doc.Elements() {
		owner, ok := e.Parent()
		if ok && e.Name().Local == element && owner.Name().Local == parent && (attribute == "" || pptxAttr(e, "", attribute) == value) {
			matches = append(matches, e)
		}
	}
	if len(matches) != 1 {
		return fmt.Errorf("added %s child count %d", element, len(matches))
	}
	if element == "Override" && pptxAttr(matches[0], "", "ContentType") != packaging.ContentTypeSlide {
		return fmt.Errorf("new slide MIME differs")
	}
	if element == "Relationship" && pptxAttr(matches[0], "", "Type") != packaging.RelTypeSlide {
		return fmt.Errorf("new slide relationship type differs")
	}
	if element == "graphicFrame" {
		if _, err := strconv.ParseUint(pptxAddedShapeID(doc, matches[0]), 10, 32); err != nil {
			return fmt.Errorf("added frame has invalid shape ID")
		}
	}
	a, b := matches[0].SourceRange()
	if a < 0 || b < a || b > len(after) {
		return fmt.Errorf("added child range invalid")
	}
	restored := append(append([]byte{}, after[:a]...), after[b:]...)
	if !bytes.Equal(before, restored) {
		return fmt.Errorf("original XML changed outside added %s", element)
	}
	return nil
}

func pptxAddedShapeID(doc *losslessxml.Document, frame losslessxml.Element) string {
	for _, e := range doc.Elements() {
		if e.Name() != pptxElem(pptxP, "cNvPr") {
			continue
		}
		p, ok := e.Parent()
		if !ok || p.Name() != pptxElem(pptxP, "nvGraphicFramePr") {
			continue
		}
		owner, ok := p.Parent()
		if ok && owner == frame {
			return pptxAttr(e, "", "id")
		}
	}
	return ""
}
func (s *pptxCaseState) validateExpected() error {
	expected, err := pptxJSON(s.record.Expected)
	if err != nil {
		return err
	}
	kind := s.record.Kind
	if kind == "reorder-refusal" {
		if s.refusal == nil {
			return fmt.Errorf("refusal absent")
		}
		_, _, _, titles, err := pptxSlideTitles(s.source)
		if err != nil {
			return err
		}
		return pptxExpectedStrings(expected, "titles", titles)
	}
	if s.output == nil {
		return fmt.Errorf("saved archive absent")
	}
	if kind == "reorder" || strings.HasPrefix(kind, "insert-") {
		paths, rids, ids, titles, err := pptxSlideTitles(s.output)
		if err != nil {
			return err
		}
		if err = pptxExpectedStrings(expected, "titles", titles); err != nil {
			return err
		}
		oldPaths, oldRIDs, oldIDs, _, err := pptxSlideTitles(s.source)
		if err != nil {
			return err
		}
		order := []int{0, 2, 1}
		if kind == "reorder" {
			for i, j := range order {
				if paths[i] != oldPaths[j] || rids[i] != oldRIDs[j] || ids[i] != oldIDs[j] {
					return fmt.Errorf("reordered slide identity differs at %d", i)
				}
			}
		} else {
			idx := 0
			if kind == "insert-middle" {
				idx = 1
			}
			if len(paths) != len(oldPaths)+1 {
				return fmt.Errorf("insert slide count")
			}
			for j := range oldPaths {
				if paths[idx] == oldPaths[j] || rids[idx] == oldRIDs[j] || ids[idx] == oldIDs[j] {
					return fmt.Errorf("inserted slide identity reused")
				}
			}
			for i, j := 0, 0; i < len(paths); i++ {
				if i == idx {
					continue
				}
				if paths[i] != oldPaths[j] || rids[i] != oldRIDs[j] || ids[i] != oldIDs[j] {
					return fmt.Errorf("existing slide identity changed at %d", i)
				}
				j++
			}
			d, err := losslessxml.Parse(s.after[paths[idx]])
			if err != nil {
				return err
			}
			sp, err := pptxShape(d, 3)
			if err != nil {
				return err
			}
			_, subtitle, err := pptxShapeParagraphs(d, sp)
			if err != nil {
				return err
			}
			if strings.Join(subtitle, "\n") != expected["insertedSubtitle"] {
				return fmt.Errorf("new subtitle %q", subtitle)
			}
		}
		return nil
	}
	if kind == "set-notes" || kind == "notes-readback" {
		d, err := losslessxml.Parse(s.after["ppt/notesSlides/notesSlide1.xml"])
		if err != nil {
			return err
		}
		var body losslessxml.Element
		count := 0
		for _, e := range d.Elements() {
			if e.Name() == pptxElem(pptxP, "ph") && pptxAttr(e, "", "type") == "body" {
				shape, ok := e.Parent()
				if !ok {
					return fmt.Errorf("missing body owner")
				}
				_ = shape
				for p, ok := e.Parent(); ok; p, ok = p.Parent() {
					if p.Name() == pptxElem(pptxP, "sp") {
						bodies := pptxChild(d, p, pptxElem(pptxP, "txBody"))
						if len(bodies) != 1 {
							return fmt.Errorf("ambiguous notes body")
						}
						body = bodies[0]
						count++
						break
					}
				}
			}
		}
		if count != 1 {
			return fmt.Errorf("notes body count %d", count)
		}
		paragraphs := pptxChild(d, body, pptxElem(pptxA, "p"))
		lines := []string{}
		for _, p := range paragraphs {
			lines = append(lines, strings.Join(pptxTexts(d, p), ""))
		}
		if err = pptxExpectedStrings(expected, "paragraphs", lines); err != nil {
			return err
		}
		if strings.Join(lines, "\n") != expected["text"] {
			return fmt.Errorf("notes text changed")
		}
		return nil
	}
	part := "ppt/slides/slide1.xml"
	id := uint32(2)
	if strings.Contains(kind, "body") || strings.HasPrefix(kind, "bullet") || kind == "clear-bullets" {
		part = "ppt/slides/slide2.xml"
		id = 4
	} else if kind == "patch-subtitle" {
		id = 3
	}
	d, err := losslessxml.Parse(s.after[part])
	if err != nil {
		return err
	}
	if strings.HasPrefix(kind, "table-") {
		return pptxCheckTable(d, expected)
	}
	sp, err := pptxShape(d, id)
	if err != nil {
		return err
	}
	paras, lines, err := pptxShapeParagraphs(d, sp)
	if err != nil {
		return err
	}
	if err = pptxExpectedStrings(expected, "paragraphs", lines); err != nil {
		return err
	}
	if want, ok := expected["text"].(string); ok && strings.Join(lines, "\n") != want {
		return fmt.Errorf("shape text=%q want=%q", strings.Join(lines, "\n"), want)
	}
	if want, ok := expected["autofit"].(string); ok {
		body := pptxChild(d, sp, pptxElem(pptxP, "txBody"))
		fit := []string{}
		for _, node := range pptxChild(d, body[0], pptxElem(pptxA, "bodyPr")) {
			for _, e := range d.Elements() {
				if p, ok := e.Parent(); ok && p == node {
					fit = append(fit, e.Name().Local)
				}
			}
		}
		if len(fit) != 1 || fit[0] != want {
			return fmt.Errorf("autofit=%v want=%s", fit, want)
		}
	}
	bullets := []map[string]any{}
	for _, para := range paras {
		props := pptxChild(d, para, pptxElem(pptxA, "pPr"))
		if len(props) == 0 {
			continue
		}
		chars := pptxChild(d, props[0], pptxElem(pptxA, "buChar"))
		if len(chars) != 1 {
			continue
		}
		lvl, e := strconv.Atoi(pptxAttr(props[0], "", "lvl"))
		if e != nil {
			return e
		}
		bullets = append(bullets, map[string]any{"character": pptxAttr(chars[0], "", "char"), "level": float64(lvl)})
	}
	if b, ok := expected["bullet"]; ok {
		if len(bullets) != 1 || !reflect.DeepEqual(bullets[0], b) {
			return fmt.Errorf("bullet=%v want=%v", bullets, b)
		}
	}
	if b, ok := expected["bullets"].([]any); ok {
		if len(bullets) != len(b) {
			return fmt.Errorf("bullets=%v want=%v", bullets, b)
		}
		for i, want := range b {
			if !reflect.DeepEqual(bullets[i], want) {
				return fmt.Errorf("bullet %d=%v want=%v", i, bullets[i], want)
			}
		}
	}
	if c, ok := expected["bulletCount"]; ok {
		if len(bullets) != int(c.(float64)) {
			return fmt.Errorf("bullet count %d", len(bullets))
		}
	}
	if runs, ok := expected["runs"].([]any); ok {
		last := paras[len(paras)-1]
		got := []map[string]any{}
		for _, r := range pptxChild(d, last, pptxElem(pptxA, "r")) {
			texts := pptxChild(d, r, pptxElem(pptxA, "t"))
			if len(texts) != 1 {
				return fmt.Errorf("run text topology")
			}
			v, _ := texts[0].Text()
			bold := false
			pr := pptxChild(d, r, pptxElem(pptxA, "rPr"))
			if len(pr) > 0 {
				bold = pptxAttr(pr[0], "", "b") == "1"
			}
			got = append(got, map[string]any{"text": v, "bold": bold})
		}
		if len(got) != len(runs) {
			return fmt.Errorf("run count %d want %d", len(got), len(runs))
		}
		for i, run := range runs {
			if !reflect.DeepEqual(got[i], run) {
				return fmt.Errorf("run %d: %v want %v", i, got[i], run)
			}
		}
	}
	return nil
}
func pptxCheckTable(d *losslessxml.Document, expected map[string]any) error {
	frames := []losslessxml.Element{}
	for _, e := range d.Elements() {
		if e.Name() == pptxElem(pptxP, "graphicFrame") {
			frames = append(frames, e)
		}
	}
	if len(frames) != 1 {
		return fmt.Errorf("table frame count %d", len(frames))
	}
	frame := frames[0]
	xforms := pptxChild(d, frame, pptxElem(pptxP, "xfrm"))
	if len(xforms) != 1 {
		return fmt.Errorf("table transform count")
	}
	for _, axis := range []struct {
		child string
		pairs [][2]string
	}{{"off", [][2]string{{"x", "x"}, {"y", "y"}}}, {"ext", [][2]string{{"cx", "width"}, {"cy", "height"}}}} {
		nodes := pptxChild(d, xforms[0], pptxElem(pptxA, axis.child))
		if len(nodes) != 1 {
			return fmt.Errorf("table %s count", axis.child)
		}
		for _, p := range axis.pairs {
			n, e := strconv.ParseInt(pptxAttr(nodes[0], "", p[0]), 10, 64)
			if e != nil {
				return e
			}
			want, e := pptxInt(expected[p[1]])
			if e != nil || n != want {
				return fmt.Errorf("table geometry %s=%d want=%d", p[1], n, want)
			}
		}
	}
	grids := []losslessxml.Element{}
	rows := []losslessxml.Element{}
	for _, e := range d.Elements() {
		if e.Name() == pptxElem(pptxA, "gridCol") && pptxDescendant(e, frame) {
			grids = append(grids, e)
		}
		if e.Name() == pptxElem(pptxA, "tr") && pptxDescendant(e, frame) {
			rows = append(rows, e)
		}
	}
	wantCols, _ := pptxInt(expected["columns"])
	wantRows, _ := pptxInt(expected["rows"])
	if len(grids) != int(wantCols) || len(rows) != int(wantRows) {
		return fmt.Errorf("table dimensions %d x %d", len(rows), len(grids))
	}
	var totalW, totalH int64
	for i, col := range grids {
		n, e := strconv.ParseInt(pptxAttr(col, "", "w"), 10, 64)
		if e != nil || n <= 0 {
			return fmt.Errorf("invalid grid width")
		}
		totalW += n
		if v, ok := expected["columnWidths"].([]any); ok {
			want, _ := pptxInt(v[i])
			if n != want {
				return fmt.Errorf("grid column width")
			}
		}
	}
	for i, row := range rows {
		n, e := strconv.ParseInt(pptxAttr(row, "", "h"), 10, 64)
		if e != nil || n <= 0 {
			return fmt.Errorf("invalid row height")
		}
		totalH += n
		if v, ok := expected["rowHeights"].([]any); ok {
			want, _ := pptxInt(v[i])
			if n != want {
				return fmt.Errorf("row height")
			}
		}
	}
	w, _ := pptxInt(expected["width"])
	h, _ := pptxInt(expected["height"])
	if totalW != w || totalH != h {
		return fmt.Errorf("table grid sums %d/%d want %d/%d", totalW, totalH, w, h)
	}
	cells, ok := expected["cells"].([]any)
	if !ok {
		return fmt.Errorf("missing cells")
	}
	for i, row := range rows {
		cs := pptxChild(d, row, pptxElem(pptxA, "tc"))
		want := cells[i].([]any)
		if len(cs) != len(want) {
			return fmt.Errorf("row %d column count", i)
		}
		for j, cell := range cs {
			got := strings.Join(pptxTexts(d, cell), "")
			if got != want[j] {
				return fmt.Errorf("cell %d,%d=%q want=%v", i, j, got, want[j])
			}
		}
	}
	return nil
}
func (s *pptxCaseState) validateCustody() error {
	if s.record.ID == "" {
		return fmt.Errorf("missing operation")
	}
	if !bytes.Equal(s.source, s.sourceCopy) {
		return fmt.Errorf("caller source modified")
	}
	if s.record.Kind == "reorder-refusal" {
		if s.refusal == nil || s.held == nil || s.held.Text() != "Alpha" {
			return fmt.Errorf("refusal or held shape missing")
		}
		return nil
	}
	if s.after == nil {
		return fmt.Errorf("missing saved archive")
	}
	allowed := map[string]bool{}
	for _, n := range s.record.ChangedMembers {
		allowed[n] = true
	}
	actualChanged := map[string]bool{}
	for n, b := range s.before {
		after, exists := s.after[n]
		if !exists {
			return fmt.Errorf("removed original member %s", n)
		}
		if bytes.Equal(b, after) {
			continue
		}
		actualChanged[n] = true
		if !allowed[n] {
			return fmt.Errorf("unrelated member %s changed", n)
		}
		if strings.HasPrefix(s.record.Kind, "insert-") && n == "ppt/_rels/presentation.xml.rels" {
			newSlide := pptxAddedSlidePart(s.before, s.after)
			if newSlide == "" {
				return fmt.Errorf("new slide absent or ambiguous")
			}
			if err := pptxAddedChildSame(b, after, "Relationship", "Relationships", "Target", strings.TrimPrefix(newSlide, "ppt/")); err != nil {
				return fmt.Errorf("%s: %w", n, err)
			}
			continue
		}
		if strings.HasPrefix(s.record.Kind, "insert-") && n == "[Content_Types].xml" {
			newSlide := pptxAddedSlidePart(s.before, s.after)
			if newSlide == "" {
				return fmt.Errorf("new slide absent or ambiguous")
			}
			if err := pptxAddedChildSame(b, after, "Override", "Types", "PartName", "/"+newSlide); err != nil {
				return fmt.Errorf("%s: %w", n, err)
			}
			continue
		}
		if strings.HasPrefix(s.record.Kind, "table-") {
			if err := pptxAddedChildSame(b, after, "graphicFrame", "spTree", "", ""); err != nil {
				return fmt.Errorf("%s: %w", n, err)
			}
			oldDoc, err := losslessxml.Parse(b)
			if err != nil {
				return err
			}
			savedDoc, err := losslessxml.Parse(after)
			if err != nil {
				return err
			}
			ids := map[string]bool{}
			for _, e := range oldDoc.Elements() {
				if e.Name() == pptxElem(pptxP, "cNvPr") {
					id := pptxAttr(e, "", "id")
					if ids[id] {
						return fmt.Errorf("source duplicate shape ID")
					}
					ids[id] = true
				}
			}
			newID := ""
			for _, e := range savedDoc.Elements() {
				if e.Name() == pptxElem(pptxP, "graphicFrame") {
					newID = pptxAddedShapeID(savedDoc, e)
				}
			}
			if newID == "" || ids[newID] {
				return fmt.Errorf("new graphic frame ID missing or reused")
			}
			continue
		}
		shapeID := uint32(2)
		if s.record.Kind == "patch-subtitle" {
			shapeID = 3
		}
		if strings.Contains(s.record.Kind, "body") || strings.HasPrefix(s.record.Kind, "bullet") || s.record.Kind == "clear-bullets" {
			shapeID = 4
		}
		if strings.HasSuffix(n, ".xml") {
			if err := pptxMaskedXMLSame(b, after, s.record.Kind, n, shapeID); err != nil {
				return fmt.Errorf("%s: %w", n, err)
			}
		}
	}
	if len(actualChanged) != len(allowed) {
		return fmt.Errorf("changed-original member count %d want %d", len(actualChanged), len(allowed))
	}
	for n := range allowed {
		if !actualChanged[n] {
			return fmt.Errorf("allowed member did not change %s", n)
		}
	}
	added := []string{}
	for n := range s.after {
		if _, ok := s.before[n]; !ok {
			added = append(added, n)
		}
	}
	if strings.HasPrefix(s.record.Kind, "insert-") {
		if len(added) != 2 {
			return fmt.Errorf("insertion added %v", added)
		}
		slide, rel := 0, 0
		for _, n := range added {
			if strings.HasPrefix(n, "ppt/slides/slide") && strings.HasSuffix(n, ".xml") {
				slide++
			} else if strings.HasPrefix(n, "ppt/slides/_rels/slide") && strings.HasSuffix(n, ".xml.rels") {
				rel++
			}
		}
		if slide != 1 || rel != 1 {
			return fmt.Errorf("insertion additions %v", added)
		}
	} else if len(added) != 0 {
		return fmt.Errorf("unexpected additions %v", added)
	}
	beforeEdges, err := pptxEdgeMap(s.source)
	if err != nil {
		return err
	}
	afterEdges, err := pptxEdgeMap(s.output)
	if err != nil {
		return err
	}
	for key, e := range beforeEdges {
		if afterEdges[key] != e {
			return fmt.Errorf("original relationship changed %s", key)
		}
	}
	if !strings.HasPrefix(s.record.Kind, "insert-") && len(afterEdges) != len(beforeEdges) {
		return fmt.Errorf("unexpected graph edge added")
	}
	if strings.HasPrefix(s.record.Kind, "insert-") {
		if len(afterEdges) != len(beforeEdges)+2 {
			return fmt.Errorf("insertion edge count differs")
		}
		newSlide := ""
		for _, n := range added {
			if strings.HasPrefix(n, "ppt/slides/slide") && strings.HasSuffix(n, ".xml") {
				newSlide = n
			}
		}
		layout := beforeEdges["ppt/slides/slide3.xml\x00rId1"].ResolvedPart
		if layout == "" || afterEdges[newSlide+"\x00rId1"].Type != packaging.RelTypeSlideLayout || afterEdges[newSlide+"\x00rId1"].ResolvedPart != layout {
			return fmt.Errorf("new slide layout relationship differs")
		}
		found := 0
		for key, e := range afterEdges {
			if _, old := beforeEdges[key]; old {
				continue
			}
			if e.Source == "ppt/presentation.xml" && e.Type == packaging.RelTypeSlide && e.ResolvedPart == newSlide {
				found++
			}
		}
		if found != 1 {
			return fmt.Errorf("new slide presentation edge absent/ambiguous")
		}
	}
	return nil
}
