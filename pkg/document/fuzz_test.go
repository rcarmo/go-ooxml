package document

import (
	"path/filepath"
	"strings"
	"testing"
)

func FuzzDocumentRoundTrip(f *testing.F) {
	for _, tc := range CommonTextCases {
		f.Add(tc.Text)
	}
	f.Add("Hello, World")
	f.Add("Tracked text")

	f.Fuzz(func(t *testing.T, text string) {
		if len(text) > 2048 {
			return
		}

		doc, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer doc.Close()

		p := doc.AddParagraph()
		p.SetText(text)
		r := p.AddRun()
		r.SetText(text)
		r.SetBold(len(text)%2 == 0)

		if len(text) > 0 {
			doc.EnableTrackChanges("Fuzzer")
			p.InsertTrackedText(text)
			_, _ = doc.Comments().Add(text, "Fuzzer", "")
		}

		path := filepath.Join(t.TempDir(), "fuzz.docx")
		if err := doc.SaveAs(path); err != nil {
			t.Fatalf("SaveAs() error = %v", err)
		}

		doc2, err := Open(path)
		if err != nil {
			t.Fatalf("Open() error = %v", err)
		}
		defer doc2.Close()

		if text != "" && len(doc2.Paragraphs()) > 0 && !strings.Contains(doc2.Paragraphs()[0].Text(), text) {
			t.Fatalf("round-trip text mismatch")
		}
	})
}

func FuzzDocumentTableOperations(f *testing.F) {
	f.Add(2, 2, "Cell")
	f.Add(3, 1, "Value")

	f.Fuzz(func(t *testing.T, rows, cols int, text string) {
		if rows < 1 || rows > 10 || cols < 1 || cols > 10 {
			return
		}
		if len(text) > 512 {
			return
		}

		doc, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer doc.Close()

		tbl := doc.AddTable(rows, cols)
		tbl.Cell(0, 0).SetText(text)

		if tbl.RowCount() != rows || tbl.ColumnCount() != cols {
			t.Fatalf("table size mismatch")
		}
	})
}

func FuzzDocumentContentControls(f *testing.F) {
	f.Add("Tag", "Alias", "Value")
	f.Add("", "", "Plain")

	f.Fuzz(func(t *testing.T, tag, alias, text string) {
		if len(tag) > 128 || len(alias) > 128 || len(text) > 1024 {
			return
		}

		doc, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer doc.Close()

		cc := doc.AddBlockContentControl(tag, alias, text)
		if cc == nil {
			t.Fatalf("AddBlockContentControl returned nil")
		}

		if tag != "" && doc.ContentControlByTag(tag) == nil {
			t.Fatalf("ContentControlByTag(%q) missing", tag)
		}
	})
}

func FuzzDocumentHyperlinks(f *testing.F) {
	f.Add("https://example.com", "Example")
	f.Add("mailto:hello@example.com", "Email")

	f.Fuzz(func(t *testing.T, url, text string) {
		if len(url) == 0 || len(url) > 256 || len(text) > 256 {
			return
		}

		doc, err := New()
		if err != nil {
			t.Fatalf("New() error = %v", err)
		}
		defer doc.Close()

		para := doc.AddParagraph()
		link, err := para.AddHyperlink(url, text)
		if err != nil {
			t.Fatalf("AddHyperlink() error = %v", err)
		}
		_ = link

		path := filepath.Join(t.TempDir(), "hyperlink.docx")
		if err := doc.SaveAs(path); err != nil {
			t.Fatalf("SaveAs() error = %v", err)
		}

		reopen, err := Open(path)
		if err != nil {
			t.Fatalf("Open() error = %v", err)
		}
		defer reopen.Close()
	})
}
