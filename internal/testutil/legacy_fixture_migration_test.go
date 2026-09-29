package testutil

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
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
	// These source rows describe the original v0.35 migration. Do not rewrite
	// their identities when a later distribution retires a physical archive.
	if ledger.Schema != 1 || ledger.ReferenceCommit != "7b7a2fa2610c421cdde9d7b1da9125c8f98b9dd8" || ledger.ReferenceTag != "v0.35.0" || identity.Tag != "v0.111.0" || len(ledger.Fixtures) != 38 || len(ledger.Specifications) != 2 {
		t.Fatal("historical migration or current pin differs from expected scope")
	}
	retired := readRetiredFixtures(t)
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
		if tombstone, retiredHere := retired[id]; retiredHere {
			if !fixture || ok || tombstone.Path != sharedPath || tombstone.Hash != hash || tombstone.Bytes != size || id != "fixture-"+hash {
				t.Fatalf("historical tombstone differs for %s", oldPath)
			}
			if _, err := os.Lstat(ReferencePath(sharedPath)); !os.IsNotExist(err) {
				t.Fatalf("retired physical archive still present: %v", err)
			}
			if _, err := LookupFixture(id); err == nil {
				t.Fatalf("retired ID still resolves: %s", id)
			}
		} else if !ok || f.path != sharedPath || f.hash != hash || f.size != size || f.role != wantRole {
			t.Fatalf("ledger differs from pinned manifest for %s", oldPath)
		}
		seenPaths[oldPath], seenIDs[id], seenNames[displayName] = true, true, true
		if _, err := os.Lstat(filepath.Join(root, filepath.FromSlash(oldPath))); !os.IsNotExist(err) {
			t.Fatalf("old binary path %s still exists: %v", oldPath, err)
		}
		if _, retiredHere := retired[id]; retiredHere {
			return // Retired bytes belong to immutable v0.40, not the current manifest.
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
	if len(retired) != 2 || !seenIDs["fixture-9726b477472ddb7595875c9f30493df2577e587416b18418d0dc7221046690fe"] {
		t.Fatal("historical retired fixture not accounted for")
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
	checkObservedGeneratedRetirement(t, byID)
}

// checkObservedGeneratedRetirement keeps the v0.43 observed-generated inputs
// recoverable through the immutable shared tag without treating historical IDs
// or labels as aliases for committed Office inputs in the current distribution.
func checkObservedGeneratedRetirement(t *testing.T, manifest map[string]struct {
	path, hash, role string
	size             int64
}) {
	t.Helper()
	var ledger struct {
		Schema int `json:"schemaVersion"`
		Source struct {
			Commit   string `json:"commit"`
			Tag      string `json:"tag"`
			Manifest string `json:"manifestSha256"`
		} `json:"releaseSource"`
		Retired []struct {
			ID     string `json:"id"`
			Path   string `json:"path"`
			Hash   string `json:"sha256"`
			Bytes  int64  `json:"bytes"`
			Role   string `json:"role"`
			Origin []struct {
				Kind string `json:"kind"`
				Path string `json:"path"`
			} `json:"origins"`
		} `json:"retired"`
	}
	data, err := os.ReadFile(ReferencePath("ledgers", "observed-generated-retirement.json"))
	if err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(data, &ledger); err != nil {
		t.Fatal(err)
	}
	if ledger.Schema != 1 || ledger.Source.Commit != "5f07417b26d199e7b6c033fcb90773ae4208e1e3" || ledger.Source.Tag != "v0.43.0" || ledger.Source.Manifest != "e970c290ce768fb9ada8f40fd83d79e68265d82c0d2131648b3cfcc0217bcfb3" || len(ledger.Retired) != 35 {
		t.Fatal("unexpected observed-generated historical release or retirement count")
	}
	retiredIDs, origins := map[string]bool{}, map[string]bool{}
	for _, row := range ledger.Retired {
		if row.ID != "fixture-"+row.Hash || len(row.Hash) != 64 || row.Role != "fixture" || row.Bytes <= 0 || !filepath.IsLocal(row.Path) || !strings.HasPrefix(row.Path, "fixtures/") || retiredIDs[row.ID] {
			t.Fatalf("invalid observed-generated retirement %s", row.ID)
		}
		retiredIDs[row.ID] = true
		if _, ok := manifest[row.ID]; ok {
			t.Fatalf("retired G identity still manifest-selected: %s", row.ID)
		}
		if _, err := os.Lstat(ReferencePath(row.Path)); !os.IsNotExist(err) {
			t.Fatalf("retired G physical path still present: %s (%v)", row.Path, err)
		}
		if _, err := LookupFixture(row.ID); err == nil {
			t.Fatalf("retired G identity still resolves: %s", row.ID)
		}
		for _, origin := range row.Origin {
			const prefix = "testdata/generated/"
			if origin.Kind != "observed-generated" || !strings.HasPrefix(origin.Path, prefix) || !filepath.IsLocal(origin.Path) || origins[origin.Path] {
				t.Fatalf("invalid or duplicate G source path %s", origin.Path)
			}
			origins[origin.Path] = true
		}
	}
	if len(origins) != 36 {
		t.Fatalf("G source paths=%d, want 36", len(origins))
	}
	for label := range fixtureIDs {
		if strings.HasPrefix(label, "generated/") {
			t.Fatalf("generated input label still active: %s", label)
		}
	}
	if len(fixtureIDs) != 38 {
		t.Fatalf("retained input labels=%d, want 38", len(fixtureIDs))
	}
}

// A retirement records old whole-archive identities without inventing a
// content-hash alias. Independently check retained member bytes against the
// sealed member inventory; v0.40 keeps the removed original archives.
type retiredFixture struct {
	ID, Path, Hash, RetainedID, RetainedPath, RetainedHash string
	Bytes                                                  int64
}

func readRetiredFixtures(t *testing.T) map[string]retiredFixture {
	t.Helper()
	// Decode exact JSON field spellings, including SHA256 fields.
	var raw struct {
		Schema int `json:"schemaVersion"`
		Source struct {
			Commit string `json:"commit"`
			Tag    string `json:"tag"`
		} `json:"releaseSource"`
		Retired []struct {
			ID           string `json:"id"`
			Path         string `json:"path"`
			Hash         string `json:"sha256"`
			Bytes        int64  `json:"bytes"`
			RetainedID   string `json:"retainedId"`
			RetainedPath string `json:"retainedPath"`
			RetainedHash string `json:"retainedSha256"`
			Members      []struct {
				Name  string `json:"name"`
				Hash  string `json:"sha256"`
				Bytes int64  `json:"bytes"`
			} `json:"members"`
		} `json:"retired"`
	}
	b, err := os.ReadFile(ReferencePath("ledgers", "fixture-content-consolidation.json"))
	if err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(b, &raw); err != nil {
		t.Fatal(err)
	}
	if raw.Schema != 1 || raw.Source.Commit != "a47c51ada71dfe1561f403c8861954bdbf5b9023" || raw.Source.Tag != "v0.40.0" || len(raw.Retired) != 2 {
		t.Fatal("unexpected historical retirement source")
	}
	result := make(map[string]retiredFixture)
	for _, r := range raw.Retired {
		if r.ID != "fixture-"+r.Hash || len(r.Hash) != 64 || !filepath.IsLocal(r.Path) || !strings.HasPrefix(r.Path, "fixtures/") || r.Bytes <= 0 || len(r.Members) == 0 || result[r.ID].ID != "" {
			t.Fatal("invalid retired fixture identity", r.ID)
		}
		wantRetained := map[string]string{
			"fixture-9726b477472ddb7595875c9f30493df2577e587416b18418d0dc7221046690fe": "fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5",
			"fixture-1780cc7a1c0ee45df6c78afcea721d995bdb67846f6bdfb78cfe78c27e28b897": "fixture-368fe96cb3ae55a0cc5fecbb599eda1d1058596d4914992f596083f291071ae4",
		}
		if wantRetained[r.ID] == "" || wantRetained[r.ID] != r.RetainedID || r.RetainedID == r.ID || r.RetainedID != "fixture-"+r.RetainedHash {
			t.Fatal("unreviewed retired-to-retained mapping", r.ID)
		}
		retainedPath, err := LookupFixture(r.RetainedID)
		if err != nil || filepath.Clean(retainedPath) != filepath.Clean(ReferencePath(r.RetainedPath)) {
			t.Fatalf("retained fixture path: %s %v", r.ID, err)
		}
		archive, err := zip.OpenReader(retainedPath)
		if err != nil {
			t.Fatal(err)
		}
		members := map[string]struct {
			hash string
			size int64
		}{}
		for _, m := range r.Members {
			if members[m.Name].hash != "" || len(m.Hash) != 64 {
				t.Fatal("invalid member record", m.Name)
			}
			members[m.Name] = struct {
				hash string
				size int64
			}{m.Hash, m.Bytes}
		}
		if len(members) != len(archive.File) {
			t.Fatal("retained member inventory differs", r.ID)
		}
		for _, f := range archive.File {
			want, ok := members[f.Name]
			if !ok {
				t.Fatal("unexpected retained member", f.Name)
			}
			reader, err := f.Open()
			if err != nil {
				t.Fatal(err)
			}
			data, readErr := io.ReadAll(reader)
			closeErr := reader.Close()
			if readErr != nil || closeErr != nil {
				t.Fatal(readErr, closeErr)
			}
			sum := sha256.Sum256(data)
			if int64(len(data)) != want.size || hex.EncodeToString(sum[:]) != want.hash {
				t.Fatal("retained member differs", f.Name)
			}
		}
		if err := archive.Close(); err != nil {
			t.Fatal(err)
		}
		item := retiredFixture{ID: r.ID, Path: r.Path, Hash: r.Hash, Bytes: r.Bytes, RetainedID: r.RetainedID, RetainedPath: r.RetainedPath, RetainedHash: r.RetainedHash}
		result[r.ID] = item
	}
	return result
}
