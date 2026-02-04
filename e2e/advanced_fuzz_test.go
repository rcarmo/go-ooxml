package e2e

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func FuzzWorkflowOpenReaderRoundTrip(f *testing.F) {
	wordFixtures := []string{
		"bullet_list.docx",
		"headers_footers.docx",
		"sdt_content_controls.docx",
		"track_changes.docx",
	}
	excelFixtures := []string{
		"comments.xlsx",
		"formulas.xlsx",
		"tables.xlsx",
	}
	pptxFixtures := []string{
		"notes.pptx",
		"tables.pptx",
		"layouts.pptx",
	}

	for _, name := range wordFixtures {
		data, err := os.ReadFile(filepath.Join("..", "testdata", "word", name))
		if err != nil {
			continue
		}
		f.Add(data, uint16(0), byte(0), uint8(0))
	}
	for _, name := range excelFixtures {
		data, err := os.ReadFile(filepath.Join("..", "testdata", "excel", name))
		if err != nil {
			continue
		}
		f.Add(data, uint16(0), byte(0), uint8(1))
	}
	for _, name := range pptxFixtures {
		data, err := os.ReadFile(filepath.Join("..", "testdata", "pptx", name))
		if err != nil {
			continue
		}
		f.Add(data, uint16(0), byte(0), uint8(2))
	}

	f.Fuzz(func(t *testing.T, data []byte, offset uint16, xor byte, kind uint8) {
		if len(data) == 0 || len(data) > 8<<20 {
			return
		}
		mutated := testutil.MutateBytes(data, offset, xor)

		switch kind % 3 {
		case 0:
			doc, err := document.OpenReader(bytes.NewReader(mutated), int64(len(mutated)))
			if err != nil {
				return
			}
			defer doc.Close()
			_ = doc.Paragraphs()
			_ = doc.Tables()
			out := filepath.Join(t.TempDir(), "roundtrip.docx")
			if err := doc.SaveAs(out); err != nil {
				return
			}
			reopen, err := document.Open(out)
			if err != nil {
				return
			}
			defer reopen.Close()
			_ = reopen.Paragraphs()
		case 1:
			wb, err := spreadsheet.OpenReader(bytes.NewReader(mutated), int64(len(mutated)))
			if err != nil {
				return
			}
			defer wb.Close()
			_ = wb.Sheets()
			out := filepath.Join(t.TempDir(), "roundtrip.xlsx")
			if err := wb.SaveAs(out); err != nil {
				return
			}
			reopen, err := spreadsheet.Open(out)
			if err != nil {
				return
			}
			defer reopen.Close()
			_ = reopen.Tables()
		default:
			pres, err := presentation.OpenReader(bytes.NewReader(mutated), int64(len(mutated)))
			if err != nil {
				return
			}
			defer pres.Close()
			_ = pres.Slides()
			out := filepath.Join(t.TempDir(), "roundtrip.pptx")
			if err := pres.SaveAs(out); err != nil {
				return
			}
			reopen, err := presentation.Open(out)
			if err != nil {
				return
			}
			defer reopen.Close()
			_ = reopen.Layouts()
		}
	})
}
