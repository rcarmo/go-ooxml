package presentation

import (
	"bytes"
	"encoding/xml"
	"reflect"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/dml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ContractPresentation is an opt-in, bounded one-title-layout creator. It
// never changes the legacy New/SaveAs template and notes-master behaviour.
// Only its owned slides are exposed through the table/cell authoring methods.
type ContractPresentation struct {
	deck   *presentationImpl
	slides []*slideImpl
	table  Table
	titles [][2]string
	grid   struct {
		rows, cols          int
		x, y, width, height int64
		cells               [][]string
	}
	closed bool
}

// NewContractPresentation narrows the embedded title template before authoring.
// Every removed layout must be owned by the template master and have exactly
// one reciprocal edge; unexpected template topology refuses before exposure.
func NewContractPresentation() (*ContractPresentation, error) {
	p, err := newFromTemplate()
	if err != nil {
		return nil, err
	}
	fail := func(reason string) (*ContractPresentation, error) {
		_ = p.Close()
		return nil, tableRefuse("PPTX_CREATION_UNSUPPORTED", reason)
	}
	if len(p.slides) != 0 || len(p.masters) != 1 || len(p.layouts) != 11 {
		return fail("unexpected template slide/master/layout inventory")
	}
	master := p.masters[0].path
	if master == "" || p.layouts[0].masterID != master {
		return fail("title layout ownership absent")
	}
	masterPart, err := p.pkg.GetPart(master)
	if err != nil {
		return fail("master absent")
	}
	source, err := masterPart.Content()
	if err != nil {
		return fail("master unreadable")
	}
	doc, err := losslessxml.Parse(source)
	if err != nil {
		return fail("master XML invalid")
	}
	wanted := map[string]bool{}
	for i, layout := range p.layouts {
		if layout.path == "" || layout.masterID != master || wanted[layout.id] {
			return fail("duplicate or foreign layout")
		}
		wanted[layout.id] = true
		owned := p.pkg.GetRelationships(master).ByID(layout.id)
		if owned == nil || owned.Type != packaging.RelTypeSlideLayout || packaging.ResolveRelationshipTarget(master, owned.Target) != layout.path || owned.TargetMode != packaging.TargetModeInternal {
			return fail("master layout edge mismatch")
		}
		reciprocal := p.pkg.GetRelationships(layout.path).ByType(packaging.RelTypeSlideMaster)
		if len(reciprocal) != 1 || packaging.ResolveRelationshipTarget(layout.path, reciprocal[0].Target) != master || reciprocal[0].TargetMode != packaging.TargetModeInternal {
			return fail("layout master edge mismatch")
		}
		if i == 0 {
			part, e := p.pkg.GetPart(layout.path)
			if e != nil {
				return fail("title layout absent")
			}
			data, e := part.Content()
			if e != nil || !bytes.Contains(data, []byte(`type="title"`)) || !bytes.Contains(data, []byte(`type="ctrTitle"`)) || !bytes.Contains(data, []byte(`type="subTitle"`)) {
				return fail("title and subtitle layout required")
			}
		}
	}
	var remove []losslessxml.Element
	seen := map[string]bool{}
	for _, node := range doc.Elements() {
		if node.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldLayoutId"}) {
			continue
		}
		parent, ok := node.Parent()
		if !ok || parent.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sldLayoutIdLst"}) {
			return fail("unexpected master layout inventory")
		}
		id := ""
		for _, a := range node.Attributes() {
			if a.Name == (xml.Name{Space: packaging.NSDocumentRelationships, Local: "id"}) {
				id = a.Value
			}
		}
		if !wanted[id] || seen[id] {
			return fail("master layout list differs from relationships")
		}
		seen[id] = true
		if id != p.layouts[0].id {
			remove = append(remove, node)
		}
	}
	if len(seen) != len(p.layouts) || len(remove) != len(p.layouts)-1 {
		return fail("master layout list incomplete")
	}
	narrowed, err := doc.RemoveElements(remove)
	if err != nil {
		return fail(err.Error())
	}
	masterPart.SetContent(narrowed)
	for _, layout := range p.layouts[1:] {
		if !p.pkg.GetRelationships(master).Remove(layout.id) {
			return fail("layout relationship removal failed")
		}
		// The relationship map is retained internally but empty owners emit no .rels.
		for _, e := range p.pkg.GetRelationships(layout.path).ByType(packaging.RelTypeSlideMaster) {
			p.pkg.GetRelationships(layout.path).Remove(e.ID)
		}
		if err := p.pkg.DeletePart(layout.path); err != nil {
			return fail("layout part removal failed")
		}
	}
	p.layouts = p.layouts[:1]
	// The printer and thumbnail are not support-graph roles. Optional
	// properties and table-style definitions retain their original identities.
	for _, v := range []struct{ source, kind, part string }{
		{packaging.PresentationPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings", "ppt/printerSettings/printerSettings1.bin"},
		{"", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail", "docProps/thumbnail.jpeg"},
	} {
		rels := p.pkg.GetRelationships(v.source).ByType(v.kind)
		if len(rels) != 1 || packaging.ResolveRelationshipTarget(v.source, rels[0].Target) != v.part {
			return fail("non-policy template support changed")
		}
		p.pkg.GetRelationships(v.source).Remove(rels[0].ID)
		if err := p.pkg.DeletePart(v.part); err != nil {
			return fail("non-policy template part removal failed")
		}
	}
	p.boundedCreation = true
	return &ContractPresentation{deck: p}, nil
}

func contractCreatorText(s string) error {
	if len(s) > 32768 || strings.ContainsAny(s, "\r\t\n") || !losslessxml.ValidAuthoredText(s) {
		return tableRefuse("PPTX_ARGUMENT_INVALID", "bounded single-line XML text required")
	}
	return nil
}

func (c *ContractPresentation) AddTitleSlide(title, subtitle string) error {
	if c == nil || c.deck == nil || c.closed {
		return tableRefuse("PPTX_CREATION_UNSUPPORTED", "missing or closed creator")
	}
	if err := contractCreatorText(title); err != nil {
		return err
	}
	if err := contractCreatorText(subtitle); err != nil {
		return err
	}
	if len(c.slides) >= 2 || c.table != nil {
		return tableRefuse("PPTX_ARGUMENT_INVALID", "two slide limit or table already authored")
	}
	s, ok := c.deck.AddSlide(0).(*slideImpl)
	if !ok {
		return tableRefuse("PPTX_CREATION_UNSUPPORTED", "slide not authored")
	}
	for i, v := range []struct{ text, kind string }{{title, "ctrTitle"}, {subtitle, "subTitle"}} {
		shape, ok := s.AddTextBox(120, int64(80+160*i), 1001, 120).(*shapeImpl)
		if !ok || shape.sp == nil {
			return tableRefuse("PPTX_CREATION_UNSUPPORTED", "placeholder not authored")
		}
		shape.sp.NvSpPr.NvPr.Ph = &dml.Ph{Type: v.kind}
		if i == 1 {
			idx := 1
			shape.sp.NvSpPr.NvPr.Ph.Idx = &idx
		}
		if err := shape.SetText(v.text); err != nil {
			return err
		}
	}
	c.slides = append(c.slides, s)
	c.titles = append(c.titles, [2]string{title, subtitle})
	return nil
}
func (c *ContractPresentation) AddTable(rows, cols int, x, y, width, height int64) (Table, error) {
	if c == nil || c.closed || c.table != nil || len(c.slides) != 1 || rows < 1 || cols < 1 || rows > 64 || cols > 64 || rows*cols > 512 || x < 0 || y < 0 || width < int64(cols) || height < int64(rows) || width > 100000000 || height > 100000000 {
		return nil, tableRefuse("PPTX_ARGUMENT_INVALID", "bounded title-table geometry required")
	}
	tbl := c.slides[0].AddTable(rows, cols, x, y, width, height)
	if tbl == nil {
		return nil, tableRefuse("PPTX_CREATION_UNSUPPORTED", "table authoring failed")
	}
	c.table = tbl
	c.grid.rows, c.grid.cols, c.grid.x, c.grid.y, c.grid.width, c.grid.height = rows, cols, x, y, width, height
	for row := 0; row < rows; row++ {
		c.grid.cells = append(c.grid.cells, make([]string, cols))
	}
	return tbl, nil
}
func (c *ContractPresentation) SlideCount() int {
	if c == nil {
		return 0
	}
	return len(c.slides)
}
func (c *ContractPresentation) SetTableCellText(table Table, row, column int, text string) error {
	if c == nil || c.deck == nil || c.closed || len(c.slides) != 1 || table == nil || row < 0 || column < 0 || row >= table.RowCount() || column >= table.ColumnCount() {
		return tableRefuse("PPTX_ARGUMENT_INVALID", "table cell is not in the created deck")
	}
	if table != c.table {
		return tableRefuse("PPTX_ARGUMENT_INVALID", "foreign table")
	}
	if err := contractCreatorText(text); err != nil {
		return err
	}
	cell := table.Cell(row, column)
	if cell == nil {
		return tableRefuse("PPTX_ARGUMENT_INVALID", "table cell absent")
	}
	cell.SetText(text)
	c.grid.cells[row][column] = text
	return nil
}

func (c *ContractPresentation) SaveAs(path string) error {
	if c == nil || c.deck == nil || c.closed || len(c.slides) == 0 {
		return tableRefuse("PPTX_CREATION_UNSUPPORTED", "title slide required")
	}
	if len(c.slides) != len(c.titles) || len(c.deck.slides) != len(c.slides) {
		return tableRefuse("PPTX_CREATION_UNSUPPORTED", "creator slide inventory changed")
	}
	for i, s := range c.slides {
		if s == nil || s.id <= 0 || s.path == "" || s.index != i || s.relID == "" || s.slide == nil || c.deck.slides[i] != s {
			return tableRefuse("PPTX_CREATION_UNSUPPORTED", "creator slide identity changed")
		}
		shapes := s.Shapes()
		wantShapes := 2
		if c.table != nil {
			wantShapes++
		}
		if len(shapes) != wantShapes {
			return tableRefuse("PPTX_CREATION_UNSUPPORTED", "creator shape inventory changed")
		}
		for j := 0; j < 2; j++ {
			if shapes[j].Text() != c.titles[i][j] {
				return tableRefuse("PPTX_CREATION_UNSUPPORTED", "creator title operand changed")
			}
		}
	}
	// Legacy SaveAs prepares the mutable package before delivery. Replay the
	// validated, bounded creator intent on a private deck so a failed save
	// cannot consume or alter this session or its held table.
	fresh, err := NewContractPresentation()
	if err != nil {
		return err
	}
	defer fresh.Close()
	for _, pair := range c.titles {
		if err := fresh.AddTitleSlide(pair[0], pair[1]); err != nil {
			return err
		}
	}
	if c.table != nil {
		if len(c.slides) != 1 || c.table.RowCount() != c.grid.rows || c.table.ColumnCount() != c.grid.cols {
			return tableRefuse("PPTX_CREATION_UNSUPPORTED", "created table topology changed")
		}
		newTable, e := fresh.AddTable(c.grid.rows, c.grid.cols, c.grid.x, c.grid.y, c.grid.width, c.grid.height)
		if e != nil {
			return e
		}
		for row := 0; row < c.grid.rows; row++ {
			for col := 0; col < c.grid.cols; col++ {
				text := c.grid.cells[row][col]
				if text != "" {
					if e := fresh.SetTableCellText(newTable, row, col, text); e != nil {
						return e
					}
				}
			}
		}
	}
	// AddTable exposes a legacy mutable Table. Reject every untracked model
	// change (including row height, frame geometry, spans and properties)
	// before SaveAs can publish a replay that silently discards that change.
	for i := range c.slides {
		if !reflect.DeepEqual(c.slides[i].slide, fresh.slides[i].slide) {
			return tableRefuse("PPTX_CREATION_UNSUPPORTED", "created slide changed outside bounded editor")
		}
	}
	return fresh.deck.SaveAs(path)
}
func (c *ContractPresentation) Close() error {
	if c == nil || c.deck == nil {
		return nil
	}
	c.closed = true
	return c.deck.Close()
}
