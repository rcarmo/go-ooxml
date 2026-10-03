package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type contract20TableGeometryWorld struct {
	deck             *presentation.ContractPresentation
	table            presentation.Table
	path             string
	archive          []byte
	members          map[string][]byte
	graph            packaging.Graph
	columns, heights []int64
	cells            [][]string
	frame            struct {
		X, Y, CX, CY int64
		ID, Name     string
	}
	slidePart string
}

func (w *contract20TableGeometryWorld) create() error {
	deck, err := presentation.NewContractPresentation()
	if err != nil {
		return err
	}
	w.deck = deck
	if err := deck.AddTitleSlide("Tables", ""); err != nil {
		return err
	}
	if deck.SlideCount() != 1 {
		return fmt.Errorf("requested slide count %d", deck.SlideCount())
	}
	return nil
}
func (w *contract20TableGeometryWorld) add() error {
	var err error
	w.table, err = w.deck.AddTable(2, 3, 120, 240, 1001, 1003)
	if err != nil {
		return err
	}
	if w.table == nil || w.table.RowCount() != 2 || w.table.ColumnCount() != 3 {
		return fmt.Errorf("table not created at requested size")
	}
	return nil
}
func (w *contract20TableGeometryWorld) text() error {
	if err := w.deck.SetTableCellText(w.table, 0, 0, "  <Alpha & Beta>  "); err != nil {
		return err
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 3; c++ {
			want := ""
			if r == 0 && c == 0 {
				want = "  <Alpha & Beta>  "
			}
			if got := w.table.Cell(r, c).Text(); got != want {
				return fmt.Errorf("cell %d,%d: %q", r, c, got)
			}
		}
	}
	return nil
}
func (w *contract20TableGeometryWorld) save() error {
	if err := w.deck.SaveAs(w.path); err != nil {
		return err
	}
	var err error
	w.archive, err = os.ReadFile(w.path)
	if err != nil {
		return err
	}
	w.members, err = blankMemberPayloads(w.archive)
	if err != nil {
		return err
	}
	w.graph, err = contract20TableGraph(w.archive)
	if err != nil {
		return err
	}
	opened, err := presentation.Open(w.path)
	if err != nil {
		return err
	}
	defer opened.Close()
	if opened.SlideCount() != 1 {
		return fmt.Errorf("reopened slide count %d", opened.SlideCount())
	}
	slide, err := opened.Slide(1)
	if err != nil {
		return err
	}
	if len(slide.Tables()) != 1 || slide.Tables()[0].RowCount() != 2 || slide.Tables()[0].ColumnCount() != 3 {
		return fmt.Errorf("reopened table dimensions")
	}
	if slide.Tables()[0].Cell(0, 0).Text() != "  <Alpha & Beta>  " {
		return fmt.Errorf("reopened first cell text")
	}
	return w.inspect()
}
func (w *contract20TableGeometryWorld) inspect() error {
	var slideParts []string
	for _, part := range w.graph.Parts {
		if part.ContentType == packaging.ContentTypeSlide {
			slideParts = append(slideParts, part.Name)
		}
	}
	if len(slideParts) != 1 {
		return fmt.Errorf("slide part count %d", len(slideParts))
	}
	w.slidePart = slideParts[0]
	var parsed struct {
		Frames []struct {
			NV struct {
				Pr struct {
					ID   string `xml:"id,attr"`
					Name string `xml:"name,attr"`
				} `xml:"cNvPr"`
				Other string `xml:"cNvGraphicFramePr"`
				NVPr  string `xml:"nvPr"`
			} `xml:"nvGraphicFramePr"`
			Xfrm struct {
				Off struct {
					X int64 `xml:"x,attr"`
					Y int64 `xml:"y,attr"`
				} `xml:"off"`
				Ext struct {
					CX int64 `xml:"cx,attr"`
					CY int64 `xml:"cy,attr"`
				} `xml:"ext"`
			} `xml:"xfrm"`
			Graphic struct {
				Data struct {
					URI   string `xml:"uri,attr"`
					Table struct {
						Grid []struct {
							W int64 `xml:"w,attr"`
						} `xml:"tblGrid>gridCol"`
						Rows []struct {
							H     int64 `xml:"h,attr"`
							Cells []struct {
								Text string `xml:"txBody>p>r>t"`
							} `xml:"tc"`
						} `xml:"tr"`
					} `xml:"tbl"`
				} `xml:"graphicData"`
			} `xml:"graphic"`
		} `xml:"cSld>spTree>graphicFrame"`
	}
	if err := xml.Unmarshal(w.members[slideParts[0]], &parsed); err != nil {
		return err
	}
	if len(parsed.Frames) != 1 {
		return fmt.Errorf("table frame count %d", len(parsed.Frames))
	}
	f := parsed.Frames[0]
	w.frame.X, w.frame.Y, w.frame.CX, w.frame.CY = f.Xfrm.Off.X, f.Xfrm.Off.Y, f.Xfrm.Ext.CX, f.Xfrm.Ext.CY
	w.frame.ID, w.frame.Name = f.NV.Pr.ID, f.NV.Pr.Name
	if f.Graphic.Data.URI != "http://schemas.openxmlformats.org/drawingml/2006/table" {
		return fmt.Errorf("wrong table graphic URI")
	}
	for _, v := range f.Graphic.Data.Table.Grid {
		w.columns = append(w.columns, v.W)
	}
	for _, r := range f.Graphic.Data.Table.Rows {
		w.heights = append(w.heights, r.H)
		cells := []string{}
		for _, c := range r.Cells {
			cells = append(cells, c.Text)
		}
		w.cells = append(w.cells, cells)
	}
	return nil
}
func (w *contract20TableGeometryWorld) dims() error {
	if len(w.heights) != 2 || len(w.columns) != 3 || w.frame.X != 120 || w.frame.Y != 240 || w.frame.CX != 1001 || w.frame.CY != 1003 {
		return fmt.Errorf("saved table dimensions %+v rows %v cols %v", w.frame, w.heights, w.columns)
	}
	return nil
}
func (w *contract20TableGeometryWorld) grid() error {
	if !reflect.DeepEqual(w.columns, []int64{333, 333, 335}) || !reflect.DeepEqual(w.heights, []int64{501, 502}) {
		return fmt.Errorf("grid %v rows %v", w.columns, w.heights)
	}
	return nil
}
func (w *contract20TableGeometryWorld) matrix() error {
	want := [][]string{{"  <Alpha & Beta>  ", "", ""}, {"", "", ""}}
	if !reflect.DeepEqual(w.cells, want) {
		return fmt.Errorf("cells %q", w.cells)
	}
	return nil
}
func (w *contract20TableGeometryWorld) titleAndProfile() error {
	if w.frame.ID == "" || w.frame.ID == "0" || w.frame.Name == "" {
		return fmt.Errorf("unnamed table or invalid frame ID %+v", w.frame)
	}
	opened, err := presentation.Open(w.path)
	if err != nil {
		return err
	}
	defer opened.Close()
	slide, err := opened.Slide(1)
	if err != nil {
		return err
	}
	shapes := slide.Shapes()
	if len(shapes) != 3 || len(slide.Tables()) != 1 || shapes[0].Text() != "Tables" || shapes[1].Text() != "" || !shapes[2].HasTable() {
		return fmt.Errorf("saved title/subtitle/table shape inventory differs: %d", len(shapes))
	}
	if err := contract20GeometryVisibleSlide(w.members[w.slidePart], "Tables", "  <Alpha & Beta>  "); err != nil {
		return err
	}
	return w.tableStructure()
}

// Inspect the complete schema-ordered direct subtree; XML unmarshal of the
// geometry/matrix above deliberately cannot certify nested property topology.
func (w *contract20TableGeometryWorld) tableStructure() error {
	doc, err := losslessxml.Parse(w.members[w.slidePart])
	if err != nil {
		return err
	}
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	children := func(parent losslessxml.Element) []losslessxml.Element {
		var result []losslessxml.Element
		for _, n := range doc.Elements() {
			owner, ok := n.Parent()
			if ok && owner == parent {
				result = append(result, n)
			}
		}
		return result
	}
	one := func(parent losslessxml.Element, kind xml.Name) (losslessxml.Element, error) {
		var result losslessxml.Element
		count := 0
		for _, n := range children(parent) {
			if n.Name() == kind {
				result = n
				count++
			}
		}
		if count != 1 {
			return result, fmt.Errorf("expected one %v, got %d", kind, count)
		}
		return result, nil
	}
	ordered := func(parent losslessxml.Element, expected ...xml.Name) error {
		got := children(parent)
		if len(got) != len(expected) {
			return fmt.Errorf("schema children of %v: %d != %d", parent.Name(), len(got), len(expected))
		}
		for i, n := range got {
			if n.Name() != expected[i] {
				return fmt.Errorf("schema child %d of %v: %v != %v", i, parent.Name(), n.Name(), expected[i])
			}
		}
		return nil
	}
	if len(doc.Elements()) == 0 || doc.Elements()[0].Name() != p("sld") {
		return fmt.Errorf("created slide root")
	}
	var frame losslessxml.Element
	frames := 0
	for _, n := range doc.Elements() {
		if n.Name() != p("graphicFrame") {
			continue
		}
		owner, ok := n.Parent()
		if ok && owner.Name() == p("spTree") {
			frame = n
			frames++
		}
	}
	if frames != 1 {
		return fmt.Errorf("direct table frame count %d", frames)
	}
	if err := ordered(frame, p("nvGraphicFramePr"), p("xfrm"), a("graphic")); err != nil {
		return err
	}
	nv, err := one(frame, p("nvGraphicFramePr"))
	if err != nil {
		return err
	}
	if err := ordered(nv, p("cNvPr"), p("cNvGraphicFramePr"), p("nvPr")); err != nil {
		return err
	}
	graphic, err := one(frame, a("graphic"))
	if err != nil {
		return err
	}
	if err := ordered(graphic, a("graphicData")); err != nil {
		return err
	}
	data, err := one(graphic, a("graphicData"))
	if err != nil {
		return err
	}
	if err := ordered(data, a("tbl")); err != nil {
		return err
	}
	table, err := one(data, a("tbl"))
	if err != nil {
		return err
	}
	got := children(table)
	if len(got) != 4 || got[0].Name() != a("tblPr") || got[1].Name() != a("tblGrid") || got[2].Name() != a("tr") || got[3].Name() != a("tr") {
		return fmt.Errorf("table property/grid/row order")
	}
	if err := ordered(got[1], a("gridCol"), a("gridCol"), a("gridCol")); err != nil {
		return err
	}
	for _, row := range got[2:] {
		if err := ordered(row, a("tc"), a("tc"), a("tc")); err != nil {
			return err
		}
		for _, cell := range children(row) {
			cellChildren := children(cell)
			if len(cellChildren) != 1 && len(cellChildren) != 2 {
				return fmt.Errorf("table cell body/property count %d", len(cellChildren))
			}
			if cellChildren[0].Name() != a("txBody") || len(cellChildren) == 2 && cellChildren[1].Name() != a("tcPr") {
				return fmt.Errorf("table cell body/property order")
			}
			body, err := one(cell, a("txBody"))
			if err != nil {
				return err
			}
			if err := ordered(body, a("bodyPr"), a("lstStyle"), a("p")); err != nil {
				return err
			}
			paragraph, err := one(body, a("p"))
			if err != nil {
				return err
			}
			runs := children(paragraph)
			if len(runs) != 1 || runs[0].Name() != a("r") {
				return fmt.Errorf("cell paragraph must have one ordinary run")
			}
			if err := ordered(runs[0], a("rPr"), a("t")); err != nil {
				// The creator may omit default rPr; a direct text leaf remains required.
				if e := ordered(runs[0], a("t")); e != nil {
					return err
				}
			}
		}
	}
	return nil
}
func (w *contract20TableGeometryWorld) graphCheck() error {
	if len(w.graph.Edges) == 0 {
		return fmt.Errorf("created package graph empty")
	}
	return contract20CreationPolicy(w.members, w.graph, 1)

}
func (w *contract20TableGeometryWorld) operandsAndFault() error {
	if !reflect.DeepEqual(w.columns, []int64{333, 333, 335}) || w.table.Cell(0, 0).Text() != "  <Alpha & Beta>  " {
		return fmt.Errorf("caller table operands changed")
	}
	before, err := blankMemberPayloads(w.archive)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(before, w.members) {
		return fmt.Errorf("saved creation members changed")
	}
	prior := filepath.Join(filepath.Dir(w.path), "prior.pptx")
	if err := os.WriteFile(prior, []byte("prior bytes"), 0600); err != nil {
		return err
	}
	failure := filepath.Join(filepath.Dir(w.path), "absent-parent", "fault.pptx")
	if err := w.deck.SaveAs(failure); err == nil {
		return fmt.Errorf("absent-parent save unexpectedly succeeded")
	}
	if _, err := os.Stat(failure); !os.IsNotExist(err) {
		return fmt.Errorf("partial fault output: %v", err)
	}
	link := filepath.Join(filepath.Dir(w.path), "existing-link.pptx")
	if err := os.Symlink(prior, link); err != nil {
		return err
	}
	if err := w.deck.SaveAs(link); err == nil {
		return fmt.Errorf("symlink save unexpectedly succeeded")
	}
	if target, err := os.Readlink(link); err != nil || target != prior {
		return fmt.Errorf("symlink destination changed: %s/%v", target, err)
	}
	b, err := os.ReadFile(prior)
	if err != nil || !bytes.Equal(b, []byte("prior bytes")) {
		return fmt.Errorf("prior destination changed: %v", err)
	}
	unchanged, err := os.ReadFile(w.path)
	if err != nil || !bytes.Equal(unchanged, w.archive) {
		return fmt.Errorf("published creation output changed: %v", err)
	}
	if w.deck.SlideCount() != 1 || w.deck.SetTableCellText(w.table, 0, 0, "bad\x00") == nil {
		return fmt.Errorf("fault consumed creator session")
	}
	if w.table.RowCount() != 2 || w.table.ColumnCount() != 3 || w.table.Cell(0, 0).Text() != "  <Alpha & Beta>  " {
		return fmt.Errorf("held table changed after save fault")
	}
	for row := 0; row < 2; row++ {
		for col := 0; col < 3; col++ {
			if row == 0 && col == 0 {
				continue
			}
			if text := w.table.Cell(row, col).Text(); text != "" {
				return fmt.Errorf("held blank cell %d,%d changed to %q", row, col, text)
			}
		}
	}
	afterPath := filepath.Join(filepath.Dir(w.path), "after-fault.pptx")
	if err := w.deck.SaveAs(afterPath); err != nil {
		return fmt.Errorf("creator session unusable after save fault: %w", err)
	}
	afterArchive, err := os.ReadFile(afterPath)
	if err != nil {
		return err
	}
	afterMembers, err := blankMemberPayloads(afterArchive)
	if err != nil {
		return err
	}
	afterGraph, err := contract20TableGraph(afterArchive)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(afterMembers, w.members) || !reflect.DeepEqual(afterGraph, w.graph) {
		return fmt.Errorf("fault changed complete regenerated creator members or graph")
	}
	entries, err := os.ReadDir(filepath.Dir(w.path))
	if err != nil {
		return err
	}
	for _, entry := range entries {
		if strings.HasPrefix(entry.Name(), ".ooxml-") {
			return fmt.Errorf("save fault left temporary entry %s", entry.Name())
		}
	}
	if _, err := os.Lstat(filepath.Join(filepath.Dir(w.path), "absent-parent")); !os.IsNotExist(err) {
		return fmt.Errorf("save fault created absent parent: %v", err)
	}
	return nil
}

func TestContract20TableGeometrySelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	const id = "@id-pptx-table-roundtrip-geometry"
	var key caseID
	var want expectedCase
	count := 0
	for k, v := range inventory {
		if k.ID == id {
			key, want = k, v
			count++
		}
	}
	if count != 1 {
		t.Fatalf("selected geometry case rows %d", count)
	}
	state := &contract20TableGeometryWorld{}
	exact := func(sc *godog.ScenarioContext, text string, action func() error) {
		sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
	}
	init := func(sc *godog.ScenarioContext) {
		sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
			*state = contract20TableGeometryWorld{path: filepath.Join(t.TempDir(), "table.pptx")}
			return ctx, nil
		})
		sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
			if state.deck != nil {
				_ = state.deck.Close()
			}
			return ctx, nil
		})
		exact(sc, `the production creator authors a new presentation with one title slide JSON "Tables"`, state.create)
		exact(sc, `the production table editor authors a 2-row 3-column table at EMU x=120 y=240 width=1001 height=1003`, state.add)
		exact(sc, `the first cell text is set to JSON "  <Alpha & Beta>  " and all other cells remain blank`, state.text)
		exact(sc, `the result is saved to a new path and independently parsed and reopened`, state.save)
		exact(sc, `the saved table has 2 rows 3 columns x=120 y=240 width=1001 height=1003`, state.dims)
		exact(sc, `the exact grid widths equal JSON [333,333,335] and exact row heights equal JSON [501,502]`, state.grid)
		exact(sc, `the exact cell matrix equals JSON [["  <Alpha & Beta>  ","",""],["","",""]]`, state.matrix)
		exact(sc, `the title remains JSON "Tables" and every table name ID namespace and schema-ordered body/property child is valid within the sealed profile`, state.titleAndProfile)
		exact(sc, `all created relationships and content types resolve inside the bounded creation support graph`, state.graphCheck)
		exact(sc, `caller geometry/text operands remain unchanged and save faults leave destination state intact`, state.operandsAndFault)
	}
	var output bytes.Buffer
	suite := godog.TestSuite{Name: "go-contract20-table-geometry", ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: id, Strict: true, Concurrency: 1}}
	if code := suite.Run(); code != 0 {
		t.Errorf("Godog %d: %s", code, output.String())
	}
	var report []reportFeature
	if err := json.Unmarshal(output.Bytes(), &report); err != nil {
		t.Fatalf("report: %v: %s", err, output.String())
	}
	data, err := json.Marshal(report)
	if err != nil {
		t.Fatal(err)
	}
	if err := reconcile(map[caseID]expectedCase{key: want}, data); err != nil {
		t.Error(err)
	}
	if dir := os.Getenv("GO_CONTRACT20_PPTX_TABLE_RECEIPTS_DIR"); dir != "" && !t.Failed() {
		if e := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(id, "@id-")+".cucumber.json"), output.Bytes(), 0600); e != nil {
			t.Fatal(e)
		}
	}
}
