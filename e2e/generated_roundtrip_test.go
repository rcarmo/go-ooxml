package e2e

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func ensureArtifactsDir(path string) error {
	return os.MkdirAll(path, 0o755)
}

func TestGeneratedRoundTripArtifacts(t *testing.T) {
	base := filepath.Join("..", "testdata", "generated")
	wordOut := filepath.Join(base, "word")
	excelOut := filepath.Join(base, "excel")
	pptxOut := filepath.Join(base, "pptx")

	for _, dir := range []string{wordOut, excelOut, pptxOut} {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			t.Fatalf("MkdirAll(%s) error = %v", dir, err)
		}
	}

	roundTripWordDir(t, filepath.Join("..", "testdata", "word"), wordOut)
	roundTripExcelDir(t, filepath.Join("..", "testdata", "excel"), excelOut)
	roundTripPptxDir(t, filepath.Join("..", "testdata", "pptx"), pptxOut)

	roundTripWordFile(t, filepath.Join("..", "testdata", "default.docx"), filepath.Join(wordOut, "default.docx"))
	roundTripPptxFile(t, filepath.Join("..", "testdata", "default.pptx"), filepath.Join(pptxOut, "default.pptx"))
}

func roundTripWordDir(t *testing.T, srcDir, dstDir string) {
	t.Helper()
	entries, err := os.ReadDir(srcDir)
	if err != nil {
		t.Fatalf("ReadDir(%s) error = %v", srcDir, err)
	}
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if !strings.HasSuffix(strings.ToLower(name), ".docx") {
			continue
		}
		t.Run("word/"+name, func(t *testing.T) {
			roundTripWordFile(t, filepath.Join(srcDir, name), filepath.Join(dstDir, name))
		})
	}
}

func roundTripExcelDir(t *testing.T, srcDir, dstDir string) {
	t.Helper()
	entries, err := os.ReadDir(srcDir)
	if err != nil {
		t.Fatalf("ReadDir(%s) error = %v", srcDir, err)
	}
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if !strings.HasSuffix(strings.ToLower(name), ".xlsx") {
			continue
		}
		t.Run("excel/"+name, func(t *testing.T) {
			roundTripExcelFile(t, filepath.Join(srcDir, name), filepath.Join(dstDir, name))
		})
	}
}

func roundTripPptxDir(t *testing.T, srcDir, dstDir string) {
	t.Helper()
	entries, err := os.ReadDir(srcDir)
	if err != nil {
		t.Fatalf("ReadDir(%s) error = %v", srcDir, err)
	}
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if !strings.HasSuffix(strings.ToLower(name), ".pptx") {
			continue
		}
		t.Run("pptx/"+name, func(t *testing.T) {
			roundTripPptxFile(t, filepath.Join(srcDir, name), filepath.Join(dstDir, name))
		})
	}
}

func roundTripWordFile(t *testing.T, srcPath, dstPath string) {
	t.Helper()
	doc, err := document.Open(srcPath)
	if err != nil {
		t.Fatalf("Open(%s) error = %v", srcPath, err)
	}
	if err := doc.SaveAs(dstPath); err != nil {
		_ = doc.Close()
		t.Fatalf("SaveAs(%s) error = %v", dstPath, err)
	}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close(%s) error = %v", srcPath, err)
	}
	round, err := document.Open(dstPath)
	if err != nil {
		t.Fatalf("Open(roundtrip %s) error = %v", dstPath, err)
	}
	if err := round.Close(); err != nil {
		t.Fatalf("Close(roundtrip %s) error = %v", dstPath, err)
	}
}

func roundTripExcelFile(t *testing.T, srcPath, dstPath string) {
	t.Helper()
	wb, err := spreadsheet.Open(srcPath)
	if err != nil {
		t.Fatalf("Open(%s) error = %v", srcPath, err)
	}
	if err := wb.SaveAs(dstPath); err != nil {
		_ = wb.Close()
		t.Fatalf("SaveAs(%s) error = %v", dstPath, err)
	}
	if err := wb.Close(); err != nil {
		t.Fatalf("Close(%s) error = %v", srcPath, err)
	}
	round, err := spreadsheet.Open(dstPath)
	if err != nil {
		t.Fatalf("Open(roundtrip %s) error = %v", dstPath, err)
	}
	if err := round.Close(); err != nil {
		t.Fatalf("Close(roundtrip %s) error = %v", dstPath, err)
	}
}

func roundTripPptxFile(t *testing.T, srcPath, dstPath string) {
	t.Helper()
	pres, err := presentation.Open(srcPath)
	if err != nil {
		t.Fatalf("Open(%s) error = %v", srcPath, err)
	}
	if err := pres.SaveAs(dstPath); err != nil {
		_ = pres.Close()
		t.Fatalf("SaveAs(%s) error = %v", dstPath, err)
	}
	if err := pres.Close(); err != nil {
		t.Fatalf("Close(%s) error = %v", srcPath, err)
	}
	round, err := presentation.Open(dstPath)
	if err != nil {
		t.Fatalf("Open(roundtrip %s) error = %v", dstPath, err)
	}
	if err := round.Close(); err != nil {
		t.Fatalf("Close(roundtrip %s) error = %v", dstPath, err)
	}
}
