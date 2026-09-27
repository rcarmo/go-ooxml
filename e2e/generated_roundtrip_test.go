package e2e

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func ensureArtifactsDir(path string) error {
	if err := testutil.CheckReferenceOutput(path); err != nil {
		return err
	}
	return os.MkdirAll(path, 0o755)
}

func TestGeneratedRoundTripArtifacts(t *testing.T) {
	base := filepath.Join("..", "artifacts", "generated")
	wordOut := filepath.Join(base, "word")
	excelOut := filepath.Join(base, "excel")
	pptxOut := filepath.Join(base, "pptx")

	for _, dir := range []string{wordOut, excelOut, pptxOut} {
		if err := ensureArtifactsDir(dir); err != nil {
			t.Fatalf("MkdirAll(%s) error = %v", dir, err)
		}
	}

	roundTripWordDir(t, "word", wordOut)
	roundTripExcelDir(t, "excel", excelOut)
	roundTripPptxDir(t, "pptx", pptxOut)

	roundTripWordFile(t, testutil.FixturePath("default.docx"), filepath.Join(wordOut, "default.docx"))
	roundTripPptxFile(t, testutil.FixturePath("default.pptx"), filepath.Join(pptxOut, "default.pptx"))
}

func roundTripWordDir(t *testing.T, prefix, dstDir string) {
	t.Helper()
	labels := testutil.FixtureLabels(prefix + "/")
	if len(labels) == 0 {
		t.Fatal("empty fixture family", prefix)
	}
	for _, label := range labels {
		name := strings.TrimPrefix(label, prefix+"/")
		if strings.Contains(name, "/") {
			continue
		}
		if !strings.HasSuffix(name, ".docx") {
			continue
		}
		t.Run(label, func(t *testing.T) { roundTripWordFile(t, testutil.FixturePath(label), filepath.Join(dstDir, name)) })
	}
}

func roundTripExcelDir(t *testing.T, prefix, dstDir string) {
	t.Helper()
	labels := testutil.FixtureLabels(prefix + "/")
	if len(labels) == 0 {
		t.Fatal("empty fixture family", prefix)
	}
	for _, label := range labels {
		name := strings.TrimPrefix(label, prefix+"/")
		if strings.Contains(name, "/") {
			continue
		}
		if !strings.HasSuffix(name, ".xlsx") {
			continue
		}
		t.Run(label, func(t *testing.T) { roundTripExcelFile(t, testutil.FixturePath(label), filepath.Join(dstDir, name)) })
	}
}

func roundTripPptxDir(t *testing.T, prefix, dstDir string) {
	t.Helper()
	labels := testutil.FixtureLabels(prefix + "/")
	if len(labels) == 0 {
		t.Fatal("empty fixture family", prefix)
	}
	for _, label := range labels {
		name := strings.TrimPrefix(label, prefix+"/")
		if strings.Contains(name, "/") {
			continue
		}
		if !strings.HasSuffix(name, ".pptx") {
			continue
		}
		t.Run(label, func(t *testing.T) { roundTripPptxFile(t, testutil.FixturePath(label), filepath.Join(dstDir, name)) })
	}
}

func roundTripWordFile(t *testing.T, srcPath, dstPath string) {
	t.Helper()
	doc, err := document.Open(srcPath)
	if err != nil {
		t.Fatalf("Open(%s) error = %v", srcPath, err)
	}
	if err := testutil.CheckReferenceOutput(dstPath); err != nil {
		t.Fatal(err)
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
	if err := testutil.CheckReferenceOutput(dstPath); err != nil {
		t.Fatal(err)
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
	if err := testutil.CheckReferenceOutput(dstPath); err != nil {
		t.Fatal(err)
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
