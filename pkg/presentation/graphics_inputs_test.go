package presentation

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

const graphicsSharedCommit = "0cf83c156c8e9433347e446e58c415db011b17be"
const graphicsManifestSHA = "05d39331f9b832b54716139e57c7f5a324a0a083eab6a12de3d61f4fe74b1967"

type graphicsAsset struct {
	ID     string `json:"id"`
	Path   string `json:"path"`
	Bytes  int    `json:"bytes"`
	SHA256 string `json:"sha256"`
	Role   string `json:"role"`
}
type graphicsOperation struct {
	Kind   string `json:"kind"`
	Part   string `json:"part"`
	Before string `json:"before"`
	After  string `json:"after"`
}
type graphicsPictureCase struct {
	ScenarioID string              `json:"scenarioId"`
	CaseID     string              `json:"caseId"`
	Operations []graphicsOperation `json:"operations"`
	Expected   []PictureInfo       `json:"expected"`
	ErrorCode  string              `json:"errorCode"`
}
type graphicsPictureRecipe struct {
	Feature       string                `json:"feature"`
	Contract      string                `json:"contract"`
	BaseFixtureID string                `json:"baseFixtureId"`
	SlidePart     string                `json:"slidePart"`
	Cases         []graphicsPictureCase `json:"cases"`
}

func graphicsRoot(t *testing.T) (string, []graphicsAsset) {
	t.Helper()
	root := os.Getenv("OOXML_GRAPHICS_ROOT")
	if root == "" {
		t.Skip("explicit sealed graphics candidate required; run make graphics-test")
	}
	root, err := filepath.Abs(root)
	if err != nil {
		t.Fatal(err)
	}
	if err := testutil.VerifyReferenceCheckout(root, testutil.ReferenceIdentity{Schema: 2, Commit: graphicsSharedCommit, Tag: "candidate-graphics", Manifest: graphicsManifestSHA}, true); err != nil {
		t.Fatal(err)
	}
	raw, err := os.ReadFile(filepath.Join(root, "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var m struct {
		Files []graphicsAsset `json:"files"`
	}
	if err = json.Unmarshal(raw, &m); err != nil {
		t.Fatal(err)
	}
	return root, m.Files
}
func graphicsReadAsset(t *testing.T, root string, assets []graphicsAsset, key string) []byte {
	t.Helper()
	for _, f := range assets {
		if f.ID != key && f.Path != key {
			continue
		}
		if !filepath.IsLocal(f.Path) {
			t.Fatal("unsafe shared asset path")
		}
		path := filepath.Join(root, f.Path)
		info, err := os.Lstat(path)
		if err != nil || !info.Mode().IsRegular() {
			t.Fatal("nonregular shared asset", path, err)
		}
		b, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		h := sha256.Sum256(b)
		if len(b) != f.Bytes || hex.EncodeToString(h[:]) != f.SHA256 {
			t.Fatal("shared asset seal differs", f.Path)
		}
		return b
	}
	t.Fatal("missing sealed shared asset", key)
	return nil
}
func graphicsMembers(t *testing.T, source []byte) map[string][]byte {
	t.Helper()
	r, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	parts := map[string][]byte{}
	for _, f := range r.File {
		if _, ok := parts[f.Name]; ok {
			t.Fatal("duplicate member")
		}
		stream, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		b, err := io.ReadAll(stream)
		closeErr := stream.Close()
		if err != nil || closeErr != nil {
			t.Fatal(err, closeErr)
		}
		parts[f.Name] = b
	}
	return parts
}
func graphicsArchive(t *testing.T, parts map[string][]byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	w := zip.NewWriter(&buf)
	names := make([]string, 0, len(parts))
	for n := range parts {
		names = append(names, n)
	}
	sort.Strings(names)
	for _, n := range names {
		entry, err := w.Create(n)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = entry.Write(parts[n]); err != nil {
			t.Fatal(err)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}
func graphicsInput(t *testing.T, source []byte, operations []graphicsOperation) []byte {
	t.Helper()
	if len(operations) == 0 {
		return bytes.Clone(source)
	}
	parts := graphicsMembers(t, source)
	for _, op := range operations {
		if op.Kind != "replace-literal-once" {
			t.Fatal("unsupported recipe", op.Kind)
		}
		b, ok := parts[op.Part]
		if !ok || strings.Count(string(b), op.Before) != 1 {
			t.Fatal("recipe literal occurrence", op.Part)
		}
		parts[op.Part] = []byte(strings.Replace(string(b), op.Before, op.After, 1))
	}
	return graphicsArchive(t, parts)
}
