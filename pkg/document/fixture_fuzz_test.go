package document

import (
	"bytes"
	"os"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func FuzzDocumentFixtureMutation(f *testing.F) {
	fixtures := []string{
		"bullet_list.docx",
		"comments.docx",
		"complex_table.docx",
		"formatted_text.docx",
		"headers_footers.docx",
		"headings.docx",
		"minimal.docx",
		"numbered_list.docx",
		"sdt_content_controls.docx",
		"simple_table.docx",
		"single_paragraph.docx",
		"styles.docx",
		"track_changes.docx",
	}

	for _, name := range fixtures {
		data, err := os.ReadFile(fixturePath(name))
		if err != nil {
			continue
		}
		f.Add(data, uint16(0), byte(0))
	}

	f.Fuzz(func(t *testing.T, data []byte, offset uint16, xor byte) {
		if len(data) == 0 || len(data) > 8<<20 {
			return
		}

		mutated := testutil.MutateBytes(data, offset, xor)
		doc, err := OpenReader(bytes.NewReader(mutated), int64(len(mutated)))
		if err != nil {
			return
		}
		defer doc.Close()

		_ = doc.Body()
		_ = doc.Paragraphs()
		_ = doc.Tables()

		path := t.TempDir() + "/fixture.docx"
		if err := doc.SaveAs(path); err != nil {
			return
		}

		reopen, err := Open(path)
		if err != nil {
			return
		}
		defer reopen.Close()

		_ = reopen.Paragraphs()
		_ = reopen.Tables()
	})
}
