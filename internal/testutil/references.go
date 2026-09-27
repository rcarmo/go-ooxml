package testutil

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
)

// ReferenceRoot selects the shared read-only reference checkout. The explicit
// override supports testing a candidate distribution before the submodule tag is
// sealed. Missing inputs must fail at the caller; there is no legacy fallback.
func ReferenceRoot() string {
	if root := os.Getenv("OOXML_FIXTURES_ROOT"); root != "" {
		return root
	}
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		panic("cannot locate shared fixture checkout")
	}
	return filepath.Join(filepath.Dir(file), "..", "..", "references", "fixtures-ooxml")
}

// ReferencePath locates a distribution input. Callers supply trusted relative
// constants, never output paths or untrusted archive-member names.
func ReferencePath(parts ...string) string {
	return filepath.Join(append([]string{ReferenceRoot()}, parts...)...)
}

// LookupFixture resolves one content ID through the only fixture manifest.
// Safe manifest-selected files are regular and their bytes match the content ID.
// Physical format/group layout belongs to the manifest, not the consumer.
func LookupFixture(id string) (string, error) {
	b, err := os.ReadFile(ReferencePath("manifest.json"))
	if err != nil {
		return "", err
	}
	var m struct {
		Schema int `json:"schemaVersion"`
		Files  []struct {
			ID     string `json:"id"`
			Path   string `json:"path"`
			Hash   string `json:"sha256"`
			Bytes  int64  `json:"bytes"`
			Role   string `json:"role"`
			Format string `json:"format"`
			Group  string `json:"scenarioGroup"`
		} `json:"files"`
	}
	if err = json.Unmarshal(b, &m); err != nil {
		return "", err
	}
	if m.Schema != 2 {
		return "", fmt.Errorf("fixture manifest schema %d, expected 2", m.Schema)
	}
	seen := map[string]bool{}
	var path, hash string
	var size int64
	for _, f := range m.Files {
		if f.Role != "fixture" {
			continue
		}
		if seen[f.ID] || f.ID != "fixture-"+f.Hash || len(f.Hash) != 64 || !filepath.IsLocal(f.Path) || strings.Contains(f.Path, "\\") || filepath.ToSlash(filepath.Clean(f.Path)) != f.Path || !strings.HasPrefix(f.Path, "fixtures/") || f.Format == "" || f.Group == "" {
			return "", fmt.Errorf("invalid/duplicate fixture identity %s", f.ID)
		}
		if _, err := hex.DecodeString(f.Hash); err != nil {
			return "", err
		}
		seen[f.ID] = true
		if f.ID == id {
			path = f.Path
			hash = f.Hash
			size = f.Bytes
		}
	}
	if path == "" {
		return "", fmt.Errorf("unknown fixture ID %s", id)
	}
	full := ReferencePath(path)
	resolved, err := filepath.EvalSymlinks(full)
	if err != nil {
		return "", err
	}
	root, err := filepath.EvalSymlinks(ReferenceRoot())
	if err != nil {
		return "", err
	}
	root, err = filepath.Abs(root)
	if err != nil {
		return "", err
	}
	resolved, err = filepath.Abs(resolved)
	if err != nil {
		return "", err
	}
	if resolved != filepath.Join(root, filepath.FromSlash(path)) || !strings.HasPrefix(resolved, root+string(filepath.Separator)) {
		return "", fmt.Errorf("symlink or escaping fixture path: %s", id)
	}
	info, err := os.Lstat(full)
	if err != nil {
		return "", err
	}
	if !info.Mode().IsRegular() {
		return "", fmt.Errorf("fixture is not regular: %s", id)
	}
	data, err := os.ReadFile(full)
	if err != nil {
		return "", err
	}
	sum := sha256.Sum256(data)
	if int64(len(data)) != size || hex.EncodeToString(sum[:]) != hash {
		return "", fmt.Errorf("fixture hash/size differs: %s", id)
	}
	return full, nil
}

// FixturePath maps trusted native test labels to canonical IDs. Labels are never
// resolved as directories and there is no legacy fixture-path fallback.
func FixturePath(parts ...string) string {
	label := strings.Join(parts, "/")
	id, ok := fixtureIDs[label]
	if !ok {
		panic("unknown native fixture label: " + label)
	}
	path, err := LookupFixture(id)
	if err != nil {
		panic(err)
	}
	return path
}

// FixtureLabels enumerates logical regression inputs, preserving provenance aliases
// without duplicating files. A prefix selects test families, not disk directories.
func FixtureLabels(prefix string) []string {
	var labels []string
	for label := range fixtureIDs {
		if strings.HasPrefix(label, prefix) {
			labels = append(labels, label)
		}
	}
	sort.Strings(labels)
	return labels
}
