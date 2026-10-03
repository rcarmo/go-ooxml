package presentation

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/pml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// These private alternates exercise the public opt-in handles and creator.
// They are native controls and confer no additional canonical case credit.
func TestContract20TableReviewAlternateBatch(t *testing.T) {
	t.Run("creator-untracked-row-height", func(t *testing.T) {
		c, err := NewContractPresentation()
		if err != nil {
			t.Fatal(err)
		}
		defer c.Close()
		if err = c.AddTitleSlide("3 < 5 & good", "Sub <&>"); err != nil {
			t.Fatal(err)
		}
		table, err := c.AddTable(3, 2, 121, 242, 1001, 1000)
		if err != nil {
			t.Fatal(err)
		}
		if err = c.SetTableCellText(table, 0, 0, "<Alpha & Beta>"); err != nil {
			t.Fatal(err)
		}
		before := filepath.Join(t.TempDir(), "before.pptx")
		if err = c.SaveAs(before); err != nil {
			t.Fatal(err)
		}
		baseline, err := os.ReadFile(before)
		if err != nil {
			t.Fatal(err)
		}
		prior := filepath.Join(t.TempDir(), "prior.pptx")
		sentinel := []byte("previous destination")
		if err = os.WriteFile(prior, sentinel, 0600); err != nil {
			t.Fatal(err)
		}
		row := table.Row(0)
		if row == nil || row.Height() != 333 {
			t.Fatalf("original row height %v", row)
		}
		row.SetHeight(334)
		if row.Height() != 334 || table.Cell(0, 0).Text() != "<Alpha & Beta>" {
			t.Fatal("altered row/text state differs")
		}
		err = c.SaveAs(prior)
		var refused *packaging.Refusal
		if !errors.As(err, &refused) || refused.Kind != "PPTX_CREATION_UNSUPPORTED" {
			t.Fatalf("untracked row mutation SaveAs = %v", err)
		}
		got, err := os.ReadFile(prior)
		if err != nil || !bytes.Equal(got, sentinel) {
			t.Fatalf("prior destination changed: %v", err)
		}
		if row.Height() != 334 || table.Cell(0, 0).Text() != "<Alpha & Beta>" {
			t.Fatal("refusal consumed held table")
		}
		row.SetHeight(333)
		after := filepath.Join(t.TempDir(), "after.pptx")
		if err = c.SaveAs(after); err != nil {
			t.Fatalf("valid creator not reusable: %v", err)
		}
		got, err = os.ReadFile(after)
		if err != nil {
			t.Fatal(err)
		}
		if !reflect.DeepEqual(contract20PPTXMembers(t, baseline), contract20PPTXMembers(t, got)) {
			t.Fatal("recovered creator members differ")
		}
	})
	for _, tc := range []struct {
		name   string
		mutate func(Table, *slideImpl)
	}{
		{"frame-geometry", func(_ Table, s *slideImpl) {
			frame := s.slide.CSld.SpTree.Content[2].(*pml.GraphicFrame)
			frame.Xfrm.Off.X++
		}},
		{"cell-span", func(table Table, _ *slideImpl) { table.Cell(1, 0).SetColSpan(2) }},
		{"table-property", func(table Table, _ *slideImpl) { table.(*tableImpl).tbl.TblPr.BandRow = new(bool) }},
	} {
		t.Run("untracked-"+tc.name, func(t *testing.T) {
			c, err := NewContractPresentation()
			if err != nil {
				t.Fatal(err)
			}
			defer c.Close()
			if err = c.AddTitleSlide("Control", "Subtitle"); err != nil {
				t.Fatal(err)
			}
			table, err := c.AddTable(2, 2, 100, 200, 1000, 1000)
			if err != nil {
				t.Fatal(err)
			}
			prior := filepath.Join(t.TempDir(), "prior.pptx")
			sentinel := []byte("prior destination")
			if err = os.WriteFile(prior, sentinel, 0600); err != nil {
				t.Fatal(err)
			}
			tc.mutate(table, c.slides[0])
			err = c.SaveAs(prior)
			var refusal *packaging.Refusal
			if !errors.As(err, &refusal) || refusal.Kind != "PPTX_CREATION_UNSUPPORTED" {
				t.Fatalf("untracked %s save = %v", tc.name, err)
			}
			got, err := os.ReadFile(prior)
			if err != nil || !bytes.Equal(got, sentinel) {
				t.Fatalf("refusal changed destination: %v", err)
			}
		})
	}
	t.Run("copied-cell-handle", func(t *testing.T) {
		source := contract20PPTXTableInput(t, "styled-table")
		before := contract20PPTXMembers(t, source)
		s, err := OpenEditing(source, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		issued, err := s.FindContractTableCell("ppt/slides/slide1.xml", 4, 0, 0)
		if err != nil {
			t.Fatal(err)
		}
		copyValue := *issued
		foreign, err := OpenEditing(source, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		foreignIssued, err := foreign.FindContractTableCell("ppt/slides/slide1.xml", 4, 0, 0)
		if err != nil {
			t.Fatal(err)
		}
		for _, tc := range []struct {
			name   string
			target *ContractTableCellTarget
		}{
			{"copy", &copyValue},
			{"forged", &ContractTableCellTarget{selected: issued.selected}},
			{"foreign", foreignIssued},
		} {
			err = s.SetContractTableCellText(tc.target, "unauthorised")
			var refused *packaging.Refusal
			if !errors.As(err, &refused) || refused.Kind != "PPTX_STALE_TABLE_HANDLE" {
				t.Fatalf("%s handle write = %v", tc.name, err)
			}
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
			t.Fatal("copied handle changed saved members")
		}
		if err = s.SetContractTableCellText(issued, "authorised"); err != nil {
			t.Fatalf("issued handle consumed by copied refusal: %v", err)
		}
		path = filepath.Join(t.TempDir(), "valid.pptx")
		if _, err = s.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		data, err = os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		got := contract20PPTXMembers(t, data)
		for part, was := range before {
			if part != "ppt/slides/slide1.xml" && !bytes.Equal(was, got[part]) {
				t.Fatalf("unrelated member changed %s", part)
			}
		}
		if bytes.Equal(before["ppt/slides/slide1.xml"], got["ppt/slides/slide1.xml"]) {
			t.Fatal("issued edit not saved")
		}
		err = s.SetContractTableCellText(issued, "again")
		var refused *packaging.Refusal
		if !errors.As(err, &refused) || refused.Kind != "PPTX_STALE_TABLE_HANDLE" {
			t.Fatalf("genuine stale generation = %v", err)
		}
	})
}
