package testutil

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestLegacyFixtureMigrationBatch keeps the human-readable test names tied to
// the exact assets, while proving that removed local paths have no fallback.
func TestLegacyFixtureMigrationBatch(t *testing.T) {
	var ledger struct {
		Schema          int    `json:"schemaVersion"`
		ReferenceCommit string `json:"referenceCommit"`
		ReferenceTag    string `json:"referenceTag"`
		Fixtures        []struct {
			OldPath       string `json:"oldPath"`
			SharedAssetID string `json:"sharedAssetId"`
			SharedPath    string `json:"sharedPath"`
			SHA256        string `json:"sha256"`
			Bytes         int64  `json:"bytes"`
			DisplayName   string `json:"displayName"`
		} `json:"fixtures"`
		Specifications []struct {
			OldPath       string `json:"oldPath"`
			SharedAssetID string `json:"sharedAssetId"`
			SharedPath    string `json:"sharedPath"`
			SHA256        string `json:"sha256"`
			Bytes         int64  `json:"bytes"`
			DisplayName   string `json:"displayName"`
		} `json:"specifications"`
	}
	b, err := os.ReadFile("../../spec/legacy-fixture-migration.json")
	if err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(b, &ledger); err != nil {
		t.Fatal(err)
	}
	pin, err := os.ReadFile("../../spec/reference-distribution.json")
	if err != nil {
		t.Fatal(err)
	}
	var identity struct {
		Commit string `json:"commit"`
		Tag    string `json:"tag"`
	}
	if err := json.Unmarshal(pin, &identity); err != nil {
		t.Fatal(err)
	}
	if ledger.Schema != 1 || ledger.ReferenceCommit != identity.Commit || ledger.ReferenceTag != identity.Tag || len(ledger.Fixtures) != 38 || len(ledger.Specifications) != 2 {
		t.Fatal("migration ledger differs from pinned reference or expected scope")
	}
	root := filepath.Clean("../..")
	manifestBytes, err := os.ReadFile(ReferencePath("manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var manifest struct {
		Files []struct {
			ID    string `json:"id"`
			Path  string `json:"path"`
			Hash  string `json:"sha256"`
			Bytes int64  `json:"bytes"`
			Role  string `json:"role"`
		} `json:"files"`
	}
	if err := json.Unmarshal(manifestBytes, &manifest); err != nil {
		t.Fatal(err)
	}
	byID := map[string]struct {
		path, hash, role string
		size             int64
	}{}
	for _, f := range manifest.Files {
		if _, exists := byID[f.ID]; exists {
			t.Fatalf("duplicate manifest ID %s", f.ID)
		}
		byID[f.ID] = struct {
			path, hash, role string
			size             int64
		}{f.Path, f.Hash, f.Role, f.Bytes}
	}
	seenPaths, seenIDs, seenNames := map[string]bool{}, map[string]bool{}, map[string]bool{}
	check := func(t *testing.T, oldPath, id, sharedPath, hash, displayName string, size int64, fixture bool) {
		t.Helper()
		if !filepath.IsLocal(oldPath) || (!strings.HasPrefix(oldPath, "testdata/") && !strings.HasPrefix(oldPath, "docs/")) ||
			!filepath.IsLocal(sharedPath) || len(hash) != 64 || size <= 0 ||
			len(strings.Fields(displayName)) < 2 || seenPaths[oldPath] || seenIDs[id] || seenNames[displayName] {
			t.Fatalf("invalid or repeated migration record %q", oldPath)
		}
		f, ok := byID[id]
		wantRole := "specification"
		if fixture {
			wantRole = "fixture"
		}
		if !ok || f.path != sharedPath || f.hash != hash || f.size != size || f.role != wantRole {
			t.Fatalf("ledger differs from pinned manifest for %s", oldPath)
		}
		seenPaths[oldPath], seenIDs[id], seenNames[displayName] = true, true, true
		if _, err := os.Lstat(filepath.Join(root, filepath.FromSlash(oldPath))); !os.IsNotExist(err) {
			t.Fatalf("old binary path %s still exists: %v", oldPath, err)
		}
		var full string
		if fixture {
			if !strings.HasPrefix(oldPath, "testdata/") || !strings.HasPrefix(sharedPath, "fixtures/") || id != "fixture-"+hash {
				t.Fatalf("wrong fixture identity %s", oldPath)
			}
			var err error
			full, err = LookupFixture(id)
			if err != nil {
				t.Fatal(err)
			}
		} else {
			if !strings.HasPrefix(oldPath, "docs/ECMA-376-") || !strings.HasPrefix(sharedPath, "specs/ecma-376/") || id != "asset-"+hash {
				t.Fatalf("wrong specification identity %s", oldPath)
			}
			full = ReferencePath(filepath.FromSlash(sharedPath))
		}
		if filepath.Clean(full) != filepath.Clean(ReferencePath(filepath.FromSlash(sharedPath))) {
			t.Fatalf("manifest path drift for %s", oldPath)
		}
		info, err := os.Lstat(full)
		if err != nil || !info.Mode().IsRegular() || info.Size() != size {
			t.Fatalf("shared asset size/mode differs for %s: %v", oldPath, err)
		}
		content, err := os.ReadFile(full)
		if err != nil {
			t.Fatal(err)
		}
		sum := sha256.Sum256(content)
		if hex.EncodeToString(sum[:]) != hash {
			t.Fatalf("shared asset bytes differ for %s", oldPath)
		}
	}
	for _, f := range ledger.Fixtures {
		f := f
		t.Run(f.DisplayName, func(t *testing.T) {
			check(t, f.OldPath, f.SharedAssetID, f.SharedPath, f.SHA256, f.DisplayName, f.Bytes, true)
		})
	}
	for _, f := range ledger.Specifications {
		f := f
		t.Run(f.DisplayName, func(t *testing.T) {
			check(t, f.OldPath, f.SharedAssetID, f.SharedPath, f.SHA256, f.DisplayName, f.Bytes, false)
		})
	}
	if len(seenPaths) != 40 || len(seenIDs) != 40 || len(seenNames) != 40 {
		t.Fatal("incomplete migration custody inventory")
	}
	remaining := 0
	if err := filepath.WalkDir("../../testdata", func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if !d.IsDir() {
			remaining++
			if filepath.ToSlash(path) != "../../testdata/FIXTURES.md" {
				return fmt.Errorf("unexpected local testdata file %s", path)
			}
		}
		return nil
	}); err != nil || remaining != 1 {
		t.Fatalf("unexpected local testdata after migration: count=%d, error=%v", remaining, err)
	}
	// The same bytes remain local only where go:embed needs them at runtime.
	var templateID string
	for _, f := range ledger.Fixtures {
		if f.OldPath == "testdata/default.pptx" {
			templateID = f.SharedAssetID
		}
	}
	if templateID == "" {
		t.Fatal("missing embedded-template source identity")
	}
	embedded, err := os.ReadFile("../../pkg/presentation/templates/default.pptx")
	if err != nil {
		t.Fatal(err)
	}
	sharedTemplate, err := LookupFixture(templateID)
	if err != nil {
		t.Fatal(err)
	}
	canonical, err := os.ReadFile(sharedTemplate)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(embedded, canonical) {
		t.Fatalf("embedded template differs from pinned shared fixture %s", templateID)
	}
}
