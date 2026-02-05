package document

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func benchmarkFixturePath(name string) string {
	_, filename, _, ok := runtime.Caller(0)
	if !ok {
		return filepath.Join("..", "..", "testdata", "word", name)
	}
	root := filepath.Join(filepath.Dir(filename), "..", "..")
	return filepath.Join(root, "testdata", "word", name)
}

func BenchmarkDocumentOpenReader(b *testing.B) {
	data, err := os.ReadFile(benchmarkFixturePath("minimal.docx"))
	if err != nil {
		b.Skipf("fixture read failed: %v", err)
	}
	b.SetBytes(int64(len(data)))
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		doc, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			b.Fatalf("OpenReader() error = %v", err)
		}
		if err := doc.Close(); err != nil {
			b.Fatalf("Close() error = %v", err)
		}
	}
}

func BenchmarkDocumentOpenReaderLarge(b *testing.B) {
	data, err := os.ReadFile(benchmarkFixturePath("headers_footers.docx"))
	if err != nil {
		b.Skipf("fixture read failed: %v", err)
	}
	b.SetBytes(int64(len(data)))
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		doc, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			b.Fatalf("OpenReader() error = %v", err)
		}
		if err := doc.Close(); err != nil {
			b.Fatalf("Close() error = %v", err)
		}
	}
}

func BenchmarkDocumentSaveAs(b *testing.B) {
	doc, err := New()
	if err != nil {
		b.Fatalf("New() error = %v", err)
	}
	defer doc.Close()
	doc.AddParagraph().SetText("Benchmark document")
	doc.AddTable(2, 2)

	tmpDir := b.TempDir()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		path := filepath.Join(tmpDir, fmt.Sprintf("bench-%d.docx", i))
		if err := doc.SaveAs(path); err != nil {
			b.Fatalf("SaveAs() error = %v", err)
		}
	}
}

func BenchmarkDocumentSaveAsLarge(b *testing.B) {
	doc, err := New()
	if err != nil {
		b.Fatalf("New() error = %v", err)
	}
	defer doc.Close()

	for i := 0; i < 200; i++ {
		para := doc.AddParagraph()
		para.SetText(fmt.Sprintf("Benchmark paragraph %d", i))
		if i%10 == 0 {
			table := doc.AddTable(3, 3)
			table.Cell(0, 0).SetText("Header")
			table.Cell(1, 1).SetText("Cell")
		}
	}

	tmpDir := b.TempDir()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		path := filepath.Join(tmpDir, fmt.Sprintf("bench-large-%d.docx", i))
		if err := doc.SaveAs(path); err != nil {
			b.Fatalf("SaveAs() error = %v", err)
		}
	}
}
