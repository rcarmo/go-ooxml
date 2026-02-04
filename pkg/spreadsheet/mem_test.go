package spreadsheet

import (
	"bytes"
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func memFixturePath(name string) string {
	_, filename, _, ok := runtime.Caller(0)
	if !ok {
		return filepath.Join("..", "..", "testdata", "excel", name)
	}
	root := filepath.Join(filepath.Dir(filename), "..", "..")
	return filepath.Join(root, "testdata", "excel", name)
}

func TestWorkbookOpenReader_MemProfile(t *testing.T) {
	if os.Getenv("ENABLE_MEMPROFILE") == "" {
		t.Skip("ENABLE_MEMPROFILE not set")
	}
	data, err := os.ReadFile(memFixturePath("minimal.xlsx"))
	if err != nil {
		t.Fatalf("read fixture: %v", err)
	}
	wb, err := OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("OpenReader() error = %v", err)
	}
	if err := wb.Close(); err != nil {
		t.Fatalf("Close() error = %v", err)
	}
}
