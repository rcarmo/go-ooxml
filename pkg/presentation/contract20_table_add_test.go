package presentation

import (
	"bytes"
	"encoding/xml"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func strconvUint(v uint32) string { return strconv.FormatUint(uint64(v), 10) }
func strconvInt(v int64) string   { return strconv.FormatInt(v, 10) }

func TestContract20RetainedTableInsertionBatch(t *testing.T) {
	for _, tc := range []struct {
		name, recipe, part  string
		rows, cols          int
		x, y, width, height int64
	}{
		{"selected", "stale-table", "ppt/slides/slide1.xml", 1, 1, 1200, 100, 900, 600},
		{"alternate", "stale-table", "ppt/slides/slide1.xml", 2, 3, 1700, 250, 1001, 1003},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := contract20PPTXTableInput(t, tc.recipe)
			original := bytes.Clone(source)
			before := contract20PPTXMembers(t, source)
			s, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			old, err := s.FindContractTableCell(tc.part, 4, 0, 0)
			if err != nil {
				t.Fatal(err)
			}
			frameID, err := s.AddContractTable(tc.part, tc.rows, tc.cols, tc.x, tc.y, tc.width, tc.height)
			if err != nil {
				t.Fatal(err)
			}
			if frameID == 0 || frameID == 4 {
				t.Errorf("new frame ID %d", frameID)
			}
			if err = s.SetContractTableCellText(old, "stale write"); err == nil {
				t.Error("old handle accepted")
			} else {
				var refused *packaging.Refusal
				if !errors.As(err, &refused) || refused.Kind != "PPTX_STALE_TABLE_HANDLE" {
					t.Errorf("old handle refusal %v", err)
				}
			}
			path := filepath.Join(t.TempDir(), "added.pptx")
			if _, err = s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			output, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			after := contract20PPTXMembers(t, output)
			if len(before) != len(after) {
				t.Errorf("member count %d != %d", len(after), len(before))
			}
			for part, oldPart := range before {
				if part != tc.part && !bytes.Equal(oldPart, after[part]) {
					t.Errorf("unselected part %s changed", part)
				}
			}
			oldXML, newXML := string(before[tc.part]), string(after[tc.part])
			if !strings.Contains(newXML, oldXML[:strings.LastIndex(oldXML, "</p:spTree>")]) {
				t.Error("original tree prefix changed")
			}
			if !strings.Contains(newXML, `<p:cNvPr id="`+strconvUint(frameID)+`"`) || !strings.Contains(newXML, `x="`+strconvInt(tc.x)+`"`) || !strings.Contains(newXML, `cx="`+strconvInt(tc.width)+`"`) {
				t.Error("new frame geometry/id absent")
			}
			var deck struct {
				Frames []struct {
					Pr struct {
						ID string `xml:"id,attr"`
					} `xml:"nvGraphicFramePr>cNvPr"`
					Table struct {
						Cols []struct {
							W int64 `xml:"w,attr"`
						} `xml:"tblGrid>gridCol"`
						Rows []struct {
							H     int64 `xml:"h,attr"`
							Cells []struct {
								Text string `xml:"txBody>p>r>t"`
							} `xml:"tc"`
						} `xml:"tr"`
					} `xml:"graphic>graphicData>tbl"`
				} `xml:"cSld>spTree>graphicFrame"`
			}
			if err := xml.Unmarshal(after[tc.part], &deck); err != nil {
				t.Fatal(err)
			}
			if len(deck.Frames) != 2 {
				t.Fatalf("frames %d", len(deck.Frames))
			}
			newFrame := deck.Frames[1]
			if newFrame.Pr.ID != strconvUint(frameID) || len(newFrame.Table.Cols) != tc.cols || len(newFrame.Table.Rows) != tc.rows {
				t.Errorf("new table shape %+v", newFrame)
			}
			reopened, err := OpenEditing(output, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if _, err := reopened.FindContractTableCell(tc.part, 4, 0, 0); err != nil {
				t.Errorf("first table missing: %v", err)
			}
			if _, err := reopened.FindContractTableCell(tc.part, frameID, 0, 0); err != nil {
				t.Errorf("new table missing: %v", err)
			}
			if !bytes.Equal(source, original) || !reflect.DeepEqual(before, contract20PPTXMembers(t, source)) {
				t.Error("caller archive mutated")
			}
		})
	}
}

func TestContract20RetainedTableInsertionRefusalBatch(t *testing.T) {
	source := contract20PPTXTableInput(t, "stale-table")
	before := contract20PPTXMembers(t, source)
	for _, tc := range []struct {
		name, part          string
		rows, cols          int
		x, y, width, height int64
	}{
		{"missing-slide", "ppt/slides/absent.xml", 1, 1, 1200, 100, 900, 600},
		{"zero-columns", "ppt/slides/slide1.xml", 1, 0, 1200, 100, 900, 600},
		{"width-too-small", "ppt/slides/slide1.xml", 1, 2, 1200, 100, 1, 600},
		{"negative-coordinate", "ppt/slides/slide1.xml", 1, 1, -1, 100, 900, 600},
	} {
		t.Run(tc.name, func(t *testing.T) {
			s, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if _, err = s.AddContractTable(tc.part, tc.rows, tc.cols, tc.x, tc.y, tc.width, tc.height); err == nil {
				t.Fatal("unsupported add accepted")
			}
			path := filepath.Join(t.TempDir(), "refused.pptx")
			if _, err = s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			data, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(before, contract20PPTXMembers(t, data)) {
				t.Error("refusal changed member payloads")
			}
		})
	}
}
