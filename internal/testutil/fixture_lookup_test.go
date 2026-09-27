package testutil

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func TestFixtureIdentityLookupBatch(t *testing.T) {
	root := t.TempDir()
	t.Setenv("OOXML_FIXTURES_ROOT", root)
	data := []byte("native fixture")
	h := sha256.Sum256(data)
	hash := hex.EncodeToString(h[:])
	id := "fixture-" + hash
	path := "fixtures/docx/native/owned.bin"
	full := filepath.Join(root, filepath.FromSlash(path))
	if err := os.MkdirAll(filepath.Dir(full), 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(full, data, 0600); err != nil {
		t.Fatal(err)
	}
	entry := map[string]any{"id": id, "path": path, "sha256": hash, "bytes": len(data), "role": "fixture", "format": "docx", "scenarioGroup": "native"}
	write := func(rows []map[string]any) {
		t.Helper()
		b, _ := json.Marshal(map[string]any{"schemaVersion": 2, "files": rows})
		if err := os.WriteFile(filepath.Join(root, "manifest.json"), b, 0600); err != nil {
			t.Fatal(err)
		}
	}
	write([]map[string]any{entry})
	got, err := LookupFixture(id)
	if err != nil || got != full {
		t.Fatalf("lookup %s %v", got, err)
	}
	for _, tc := range []struct {
		name, key string
		value     any
	}{{"traversal", "path", "../escape"}, {"nonfixtures", "path", "metadata/x"}, {"noncanonical", "path", "fixtures/a/../b"}, {"backslash", "path", `fixtures\a`}, {"wrong identity", "id", "fixture-bad"}, {"wrong size", "bytes", 1}, {"missing format", "format", ""}, {"missing group", "scenarioGroup", ""}} {
		t.Run(tc.name, func(t *testing.T) {
			copy := map[string]any{}
			for k, v := range entry {
				copy[k] = v
			}
			copy[tc.key] = tc.value
			write([]map[string]any{copy})
			if _, err := LookupFixture(id); err == nil {
				t.Fatal("invalid fixture accepted")
			}
		})
	}
	write([]map[string]any{entry, entry})
	if _, err := LookupFixture(id); err == nil {
		t.Fatal("duplicate accepted")
	}
	write([]map[string]any{entry})
	if _, err := LookupFixture("fixture-missing"); err == nil {
		t.Fatal("missing ID accepted")
	}
	if err := os.WriteFile(full, []byte("altered"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := LookupFixture(id); err == nil {
		t.Fatal("altered bytes accepted")
	}
}
