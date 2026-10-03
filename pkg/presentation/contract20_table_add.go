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

// AddContractTable inserts one ordinary, empty rectangular table into an
// enrolled slide. Geometry, dimensions and slide part are caller-selected;
// the frame ID is allocated from the observed object-ID inventory.
// Existing member payloads and the held source archive are not rebuilt.
func (s *EditSession) AddContractTable(part string, rows, columns int, x, y, width, height int64) (uint32, error) {
	if !s.slides[part] {
		return 0, tableRefuse("missing_target", "unenrolled slide")
	}
	if rows < 1 || rows > 64 || columns < 1 || columns > 64 || int64(rows)*int64(columns) > 512 || x < 0 || y < 0 || width < int64(columns) || height < int64(rows) || width > 100000000 || height > 100000000 || x > 100000000 || y > 100000000 {
		return 0, tableRefuse("PPTX_ARGUMENT_INVALID", "bounded table geometry required")
	}
	if err := s.manipulationProtection(); err != nil {
		return 0, err
	}
	source, hash, err := s.pkg.Part(part)
	if err != nil {
		return 0, err
	}
	doc, err := losslessxml.Parse(source)
	if err != nil {
		return 0, err
	}
	es := doc.Elements()
	if len(es) == 0 || es[0].Name() != name(packaging.NSPresentationML, "sld") {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "slide root differs")
	}
	var tree losslessxml.Element
	treeCount := 0
	for _, node := range es {
		if node.Name() != name(packaging.NSPresentationML, "spTree") {
			continue
		}
		parent, ok := node.Parent()
		if !ok || parent.Name() != name(packaging.NSPresentationML, "cSld") {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "shape tree owner differs")
		}
		owner, ok := parent.Parent()
		if !ok || owner != es[0] {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "slide owner differs")
		}
		tree = node
		treeCount++
	}
	if treeCount != 1 || tree.SelfClosing() {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "one ordinary shape tree required")
	}
	// Unknown direct objects are not rewritten, but object identity must remain
	// unambiguous. Reject duplicates and unsupported group nesting before write.
	ids := map[uint32]bool{}
	for _, node := range es {
		if node.Name() != name(packaging.NSPresentationML, "cNvPr") {
			continue
		}
		owner, ok := node.Parent()
		if !ok {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "object owner absent")
		}
		frame, ok := owner.Parent()
		if !ok {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "object frame absent")
		}
		if frame != tree { // cNvPr of nvGrpSpPr belongs directly to the shape tree.
			container, ok := frame.Parent()
			if !ok || container != tree {
				return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "nested object cannot be reindexed")
			}
		} else if owner.Name() != name(packaging.NSPresentationML, "nvGrpSpPr") {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "unexpected group root identity")
		}
		var idText string
		for _, a := range node.Attributes() {
			if a.Name.Space == "" && a.Name.Local == "id" {
				idText = a.Value
			}
		}
		id, parseErr := strconv.ParseUint(idText, 10, 32)
		if parseErr != nil || id == 0 || ids[uint32(id)] {
			return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "invalid or duplicate object ID")
		}
		ids[uint32(id)] = true
	}
	frameID := uint32(0)
	for id := uint32(1); id < 100000; id++ {
		if !ids[id] {
			frameID = id
			break
		}
	}
	if frameID == 0 {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "object ID space exhausted")
	}
	// The source's DrawingML namespace must be bound in the target tree. A
	// structured insertion uses expanded names and cannot retarget a prefix.
	ns := tree.Namespaces()
	pPrefix, aPrefix := "", ""
	for prefix, uri := range ns {
		if uri == packaging.NSPresentationML && pPrefix == "" {
			pPrefix = prefix
		}
		if uri == packaging.NSDrawingML && aPrefix == "" {
			aPrefix = prefix
		}
	}
	if pPrefix == "" || aPrefix == "" {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "table prefixes not in scope")
	}
	pn := func(local string) xml.Name { return name(packaging.NSPresentationML, local) }
	an := func(local string) xml.Name { return name(packaging.NSDrawingML, local) }
	attr := func(key, value string) xml.Attr { return xml.Attr{Name: xml.Name{Local: key}, Value: value} }
	colWidth, rowHeight := width/int64(columns), height/int64(rows)
	grid := losslessxml.NewElement{Name: an("tblGrid")}
	for i := 0; i < columns; i++ {
		v := colWidth
		if i == columns-1 {
			v = width - colWidth*int64(columns-1)
		}
		grid.Children = append(grid.Children, losslessxml.NewElement{Name: an("gridCol"), Attributes: []xml.Attr{attr("w", strconv.FormatInt(v, 10))}})
	}
	table := losslessxml.NewElement{Name: an("tbl"), Children: []losslessxml.NewElement{{Name: an("tblPr")}, grid}}
	for r := 0; r < rows; r++ {
		rh := rowHeight
		if r == rows-1 {
			rh = height - rowHeight*int64(rows-1)
		}
		row := losslessxml.NewElement{Name: an("tr"), Attributes: []xml.Attr{attr("h", strconv.FormatInt(rh, 10))}}
		for c := 0; c < columns; c++ {
			row.Children = append(row.Children, losslessxml.NewElement{Name: an("tc"), Children: []losslessxml.NewElement{
				{Name: an("txBody"), Children: []losslessxml.NewElement{{Name: an("bodyPr")}, {Name: an("lstStyle")}, {Name: an("p")}}},
				{Name: an("tcPr")},
			}})
		}
		table.Children = append(table.Children, row)
	}
	frame := losslessxml.NewElement{Name: pn("graphicFrame"), Children: []losslessxml.NewElement{
		{Name: pn("nvGraphicFramePr"), Children: []losslessxml.NewElement{
			{Name: pn("cNvPr"), Attributes: []xml.Attr{attr("id", strconv.FormatUint(uint64(frameID), 10)), attr("name", fmt.Sprintf("Table %d", frameID))}},
			{Name: pn("cNvGraphicFramePr")}, {Name: pn("nvPr")},
		}},
		{Name: pn("xfrm"), Children: []losslessxml.NewElement{
			{Name: an("off"), Attributes: []xml.Attr{attr("x", strconv.FormatInt(x, 10)), attr("y", strconv.FormatInt(y, 10))}},
			{Name: an("ext"), Attributes: []xml.Attr{attr("cx", strconv.FormatInt(width, 10)), attr("cy", strconv.FormatInt(height, 10))}},
		}},
		{Name: an("graphic"), Children: []losslessxml.NewElement{{Name: an("graphicData"), Attributes: []xml.Attr{attr("uri", "http://schemas.openxmlformats.org/drawingml/2006/table")}, Children: []losslessxml.NewElement{table}}}},
	}}
	data, err := doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: tree, Children: []losslessxml.NewElement{frame}}})
	if err != nil {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", err.Error())
	}
	if bytes.Equal(source, data) || !strings.Contains(string(data), "graphicFrame") {
		return 0, tableRefuse("PPTX_TABLE_STRUCTURE_UNSUPPORTED", "table insertion missing")
	}
	if err = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: data}}); err != nil {
		return 0, err
	}
	s.generation++
	return frameID, nil
}
