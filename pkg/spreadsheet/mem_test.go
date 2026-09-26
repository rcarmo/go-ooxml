package spreadsheet

import (
	"bytes"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"testing"
)

func memFixturePath(name string) string { return testutil.FixturePath("excel", name) }

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
