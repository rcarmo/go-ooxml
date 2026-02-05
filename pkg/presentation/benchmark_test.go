package presentation

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
		return filepath.Join("..", "..", "testdata", "pptx", name)
	}
	root := filepath.Join(filepath.Dir(filename), "..", "..")
	return filepath.Join(root, "testdata", "pptx", name)
}

func BenchmarkPresentationOpenReader(b *testing.B) {
	data, err := os.ReadFile(benchmarkFixturePath("minimal.pptx"))
	if err != nil {
		b.Skipf("fixture read failed: %v", err)
	}
	b.SetBytes(int64(len(data)))
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		pres, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			b.Fatalf("OpenReader() error = %v", err)
		}
		if err := pres.Close(); err != nil {
			b.Fatalf("Close() error = %v", err)
		}
	}
}

func BenchmarkPresentationOpenReaderLarge(b *testing.B) {
	data, err := os.ReadFile(benchmarkFixturePath("images.pptx"))
	if err != nil {
		b.Skipf("fixture read failed: %v", err)
	}
	b.SetBytes(int64(len(data)))
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		pres, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			b.Fatalf("OpenReader() error = %v", err)
		}
		if err := pres.Close(); err != nil {
			b.Fatalf("Close() error = %v", err)
		}
	}
}

func BenchmarkPresentationSaveAs(b *testing.B) {
	pres, err := New()
	if err != nil {
		b.Fatalf("New() error = %v", err)
	}
	defer pres.Close()
	slide := pres.AddSlide(0)
	slide.AddTextBox(0, 0, 4000000, 1000000).SetText("Benchmark")

	tmpDir := b.TempDir()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		path := filepath.Join(tmpDir, fmt.Sprintf("bench-%d.pptx", i))
		if err := pres.SaveAs(path); err != nil {
			b.Fatalf("SaveAs() error = %v", err)
		}
	}
}

func BenchmarkPresentationSaveAsLarge(b *testing.B) {
	pres, err := New()
	if err != nil {
		b.Fatalf("New() error = %v", err)
	}
	defer pres.Close()

	for i := 0; i < 20; i++ {
		slide := pres.AddSlide(0)
		slide.AddTextBox(0, 0, 4000000, 1000000).SetText(fmt.Sprintf("Benchmark slide %d", i))
		table := slide.AddTable(2, 2, 100000, 100000, 3000000, 1000000)
		table.Cell(0, 0).SetText("Header")
	}

	tmpDir := b.TempDir()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		path := filepath.Join(tmpDir, fmt.Sprintf("bench-large-%d.pptx", i))
		if err := pres.SaveAs(path); err != nil {
			b.Fatalf("SaveAs() error = %v", err)
		}
	}
}
