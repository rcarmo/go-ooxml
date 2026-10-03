package presentation

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

func TestContract20CreatedTableGeometry(t *testing.T) {
	p, err := New()
	if err != nil {
		t.Fatal(err)
	}
	defer p.Close()
	slide := p.AddSlide(0)
	if slide == nil {
		t.Fatal("missing slide")
	}
	title := slide.AddTextBox(120, 80, 1001, 120)
	if title == nil {
		t.Fatal("title textbox missing")
	}
	title.SetText("Tables")
	table := slide.AddTable(2, 3, 120, 240, 1001, 1003)
	if table == nil || table.RowCount() != 2 || table.ColumnCount() != 3 {
		t.Fatal("table dimensions")
	}
	if cell := table.Cell(0, 0); cell == nil {
		t.Fatal("first cell missing")
	} else {
		cell.SetText("  <Alpha & Beta>  ")
	}
	path := filepath.Join(t.TempDir(), "table.pptx")
	if err := p.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	bytes, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	z, err := zip.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	defer z.Close()
	var part []byte
	for _, file := range z.File {
		if file.Name == "ppt/slides/slide1.xml" {
			r, e := file.Open()
			if e != nil {
				t.Fatal(e)
			}
			part, e = io.ReadAll(r)
			closeErr := r.Close()
			if e != nil || closeErr != nil {
				t.Fatal(e, closeErr)
			}
		}
	}
	if len(part) == 0 {
		t.Fatal("slide part absent")
	}
	var deck struct {
		Frames []struct {
			Xfrm struct {
				Off struct {
					X, Y int64 `xml:",attr"`
				} `xml:"off"`
				Ext struct {
					CX int64 `xml:"cx,attr"`
					CY int64 `xml:"cy,attr"`
				} `xml:"ext"`
			} `xml:"xfrm"`
			Graphic struct {
				Data struct {
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
	if err := xml.Unmarshal(part, &deck); err != nil {
		t.Fatal(err)
	}
	if len(deck.Frames) != 1 {
		t.Fatalf("table frames %d", len(deck.Frames))
	}
	frame := deck.Frames[0]
	var columns, rows []int64
	for _, grid := range frame.Graphic.Data.Table.Grid {
		columns = append(columns, grid.W)
	}
	for _, row := range frame.Graphic.Data.Table.Rows {
		rows = append(rows, row.H)
	}
	if !reflect.DeepEqual(columns, []int64{333, 333, 335}) || !reflect.DeepEqual(rows, []int64{501, 502}) {
		t.Errorf("table geometry columns %v rows %v", columns, rows)
	}
	if len(frame.Graphic.Data.Table.Rows) == 2 && len(frame.Graphic.Data.Table.Rows[0].Cells) == 3 {
		if text := frame.Graphic.Data.Table.Rows[0].Cells[0].Text; text != "  <Alpha & Beta>  " {
			t.Errorf("table first text %q", text)
		}
	} else {
		t.Error("table cells missing")
	}
	if len(bytes) == 0 {
		t.Error("empty saved package")
	}
	reopened, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	first, err := reopened.Slide(1)
	if err != nil {
		t.Fatal(err)
	}
	if len(first.Shapes()) == 0 || first.Shapes()[0].Text() != "Tables" {
		t.Errorf("reopened first slide title textbox missing")
	}
}
