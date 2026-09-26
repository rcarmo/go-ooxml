package acceptance

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/cucumber/gherkin/go/v26"
	"github.com/rcarmo/go-ooxml/internal/testutil"
)

type referencePin struct {
	Schema    int    `json:"schema"`
	Commit    string `json:"commit"`
	Manifest  string `json:"manifest_sha256"`
	Pack      string `json:"shared_pack_sha256"`
	Assets    int    `json:"assets"`
	Facts     int    `json:"facts"`
	Workflows int    `json:"workflows"`
	Cases     int    `json:"workflow_cases"`
}

func loadReferencePin(t *testing.T) referencePin {
	t.Helper()
	b, err := os.ReadFile("../spec/reference-distribution.json")
	if err != nil {
		t.Fatal(err)
	}
	var p referencePin
	if err = json.Unmarshal(b, &p); err != nil {
		t.Fatal(err)
	}
	if p.Schema != 1 || len(p.Commit) != 40 || p.Assets <= 0 || p.Workflows <= 0 || p.Cases <= 0 {
		t.Fatal("invalid reference pin")
	}
	return p
}

func TestPinnedReferenceDistribution(t *testing.T) {
	pin := loadReferencePin(t)
	b, err := pinnedFile(testutil.ReferenceRoot(), "manifest.json", pin.Manifest)
	if err != nil {
		t.Fatal(err)
	}
	var manifest struct {
		Schema int `json:"schemaVersion"`
		Files  []struct {
			Path   string          `json:"path"`
			Bytes  int64           `json:"bytes"`
			Hash   string          `json:"sha256"`
			Role   string          `json:"role"`
			Origin json.RawMessage `json:"origin"`
		} `json:"files"`
	}
	if err = json.Unmarshal(b, &manifest); err != nil {
		t.Fatal(err)
	}
	if manifest.Schema != 1 || len(manifest.Files) != pin.Assets {
		t.Fatal("distribution inventory differs")
	}
	seen := map[string]bool{}
	for _, f := range manifest.Files {
		if !filepath.IsLocal(f.Path) || filepath.ToSlash(filepath.Clean(f.Path)) != f.Path || strings.Contains(f.Path, "\\") || seen[f.Path] || f.Bytes < 0 || len(f.Hash) != 64 || f.Role == "" || len(f.Origin) < 3 {
			t.Fatalf("invalid manifest entry %s", f.Path)
		}
		seen[f.Path] = true
		path := testutil.ReferencePath(f.Path)
		info, err := os.Lstat(path)
		if err != nil {
			t.Fatal(err)
		}
		if !info.Mode().IsRegular() {
			t.Fatal("nonregular reference asset", f.Path)
		}
		file, err := os.Open(path)
		if err != nil {
			t.Fatal(err)
		}
		h := sha256.New()
		n, err := io.Copy(h, file)
		closeErr := file.Close()
		if err != nil || closeErr != nil {
			t.Fatal(err, closeErr)
		}
		if n != f.Bytes || hex.EncodeToString(h.Sum(nil)) != f.Hash {
			t.Fatal("reference asset hash/size mismatch", f.Path)
		}
	}
	for _, area := range []string{"fixtures", "reference-assets", "shared/v2/pack"} {
		if err := filepath.WalkDir(testutil.ReferencePath(area), func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if d.IsDir() {
				return nil
			}
			rel, err := filepath.Rel(testutil.ReferenceRoot(), path)
			if err != nil {
				return err
			}
			rel = filepath.ToSlash(rel)
			if rel == "shared/v2/pack/pack-manifest.json" {
				return nil
			}
			if !seen[rel] {
				return fmt.Errorf("unmanifested reference asset %s", rel)
			}
			return nil
		}); err != nil {
			t.Fatal(err)
		}
	}
	if pin.Pack != sharedPackHash {
		t.Fatal("pack constant and pin differ")
	}
	t.Logf("%d distribution assets independently verified", len(seen))
}

// Compile shared features only for exact inventory reconciliation. No steps run
// here, and no consumer receives execution credit from another language's state.
func TestSharedWorkflowCaseInventory(t *testing.T) {
	pin := loadReferencePin(t)
	var ledger struct {
		Features  []string `json:"features"`
		Workflows []struct {
			ID      string `json:"id"`
			Feature string `json:"feature"`
			Cases   int    `json:"expandedCases"`
		} `json:"workflows"`
	}
	if err := readReferenceJSON("ledgers/workflows.json", &ledger); err != nil {
		t.Fatal(err)
	}
	type record struct {
		feature string
		cases   int
	}
	compiled := map[string]record{}
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	total := 0
	for _, path := range ledger.Features {
		f, err := os.Open(testutil.ReferencePath(path))
		if err != nil {
			t.Fatal(err)
		}
		doc, err := gherkin.ParseGherkinDocument(f, next)
		_ = f.Close()
		if err != nil {
			t.Fatal(err)
		}
		for _, p := range gherkin.Pickles(*doc, path, next) {
			id := ""
			for _, tag := range p.Tags {
				if strings.HasPrefix(tag.Name, "@id-") {
					if id != "" {
						t.Fatal("multiple workflow IDs", path)
					}
					id = tag.Name
				}
			}
			if id == "" {
				t.Fatal("missing workflow ID", path)
			}
			r := compiled[id]
			if r.feature != "" && r.feature != path {
				t.Fatal("workflow ID reused across features", id)
			}
			r.feature = path
			r.cases++
			compiled[id] = r
			total++
		}
	}
	if len(compiled) != pin.Workflows || total != pin.Cases || len(ledger.Workflows) != pin.Workflows {
		t.Fatalf("workflow inventory %d/%d differs from pin %d/%d", len(compiled), total, pin.Workflows, pin.Cases)
	}
	seen := map[string]bool{}
	for _, w := range ledger.Workflows {
		r, ok := compiled[w.ID]
		if !ok || seen[w.ID] || r.feature != w.Feature || r.cases != w.Cases {
			t.Fatalf("workflow mismatch %+v", w)
		}
		seen[w.ID] = true
	}
	facts := 0
	for _, name := range []string{"content-types", "namespaces", "relationships", "constants"} {
		var r struct {
			Values []referenceFact `json:"values"`
		}
		if err := readReferenceJSON("facts/"+name+".json", &r); err != nil {
			t.Fatal(err)
		}
		facts += len(r.Values)
	}
	if facts != pin.Facts {
		t.Fatal("fact inventory differs", facts, pin.Facts)
	}
}
