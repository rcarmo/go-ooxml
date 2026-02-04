package spreadsheet

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
		return filepath.Join("..", "..", "testdata", "excel", name)
	}
	root := filepath.Join(filepath.Dir(filename), "..", "..")
	return filepath.Join(root, "testdata", "excel", name)
}

func BenchmarkWorkbookOpenReader(b *testing.B) {
	data, err := os.ReadFile(benchmarkFixturePath("minimal.xlsx"))
	if err != nil {
		b.Skipf("fixture read failed: %v", err)
	}
	b.SetBytes(int64(len(data)))
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		wb, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			b.Fatalf("OpenReader() error = %v", err)
		}
		if err := wb.Close(); err != nil {
			b.Fatalf("Close() error = %v", err)
		}
	}
}

func BenchmarkWorkbookSaveAs(b *testing.B) {
	wb, err := New()
	if err != nil {
		b.Fatalf("New() error = %v", err)
	}
	defer wb.Close()
	sheet := wb.SheetsRaw()[0]
	sheet.Cell("A1").SetValue("Benchmark")
	sheet.Cell("B2").SetValue(123.45)

	tmpDir := b.TempDir()
	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		path := filepath.Join(tmpDir, fmt.Sprintf("bench-%d.xlsx", i))
		if err := wb.SaveAs(path); err != nil {
			b.Fatalf("SaveAs() error = %v", err)
		}
	}
}
