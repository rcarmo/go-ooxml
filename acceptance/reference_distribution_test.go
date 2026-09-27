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
	Tag       string `json:"tag"`
	TagObject string `json:"tag_object"`
	Manifest  string `json:"manifest_sha256"`
	Pack      string `json:"shared_pack_sha256"`
	Assets    int    `json:"assets"`
	Facts     int    `json:"facts"`
	Workflows int    `json:"workflows"`
	Cases     int    `json:"workflow_cases"`
}

func loadReferencePin(t *testing.T) referencePin {
	t.Helper()
	path := os.Getenv("OOXML_REFERENCE_PIN")
	if path == "" {
		path = "../spec/reference-distribution.json"
	}
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	var p referencePin
	if err = json.Unmarshal(b, &p); err != nil {
		t.Fatal(err)
	}
	if (p.Schema != 1 && p.Schema != 2) || len(p.Commit) != 40 || p.Assets <= 0 || p.Workflows <= 0 || p.Cases <= 0 {
		t.Fatal("invalid reference pin")
	}
	candidate := os.Getenv("OOXML_REFERENCE_PIN") != "" && p.TagObject == "" && strings.HasPrefix(p.Tag, "candidate-")
	if candidate && os.Getenv("OOXML_FIXTURES_ROOT") == "" {
		t.Fatal("candidate reference pin requires explicit candidate root")
	}
	if err := testutil.VerifyReferenceCheckout(testutil.ReferenceRoot(), testutil.ReferenceIdentity{Schema: p.Schema, Commit: p.Commit, Tag: p.Tag, TagObject: p.TagObject, Manifest: p.Manifest, Pack: p.Pack}, candidate); err != nil {
		t.Fatal(err)
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
			ID     string          `json:"id"`
			Format string          `json:"format"`
			Group  string          `json:"scenarioGroup"`
			Path   string          `json:"path"`
			Bytes  int64           `json:"bytes"`
			Hash   string          `json:"sha256"`
			Role   string          `json:"role"`
			Origin json.RawMessage `json:"origins"`
		} `json:"files"`
	}
	if err = json.Unmarshal(b, &manifest); err != nil {
		t.Fatal(err)
	}
	if manifest.Schema != 2 || len(manifest.Files) != pin.Assets {
		t.Fatal("distribution inventory differs")
	}
	seen := map[string]bool{}
	ids := map[string]bool{}
	fixtureHashes := map[string]bool{}
	for _, f := range manifest.Files {
		if !filepath.IsLocal(f.Path) || filepath.ToSlash(filepath.Clean(f.Path)) != f.Path || strings.Contains(f.Path, "\\") || seen[f.Path] || f.Bytes < 0 || len(f.Hash) != 64 || f.Role == "" || len(f.Origin) < 3 {
			t.Fatalf("invalid manifest entry %s", f.Path)
		}
		if f.ID == "" || ids[f.ID] {
			t.Fatal("invalid/duplicate asset ID", f.ID)
		}
		ids[f.ID] = true
		if f.Role == "fixture" {
			if fixtureHashes[f.Hash] || f.ID != "fixture-"+f.Hash || !strings.HasPrefix(f.Path, "fixtures/") || f.Format == "" || f.Group == "" {
				t.Fatal("duplicate or invalid fixture", f.ID)
			}
			fixtureHashes[f.Hash] = true
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
	areas := []string{"fixtures", "notices"}
	if pin.Schema == 1 {
		areas = append(areas, "shared/v2/pack")
	} else {
		const contractPath = "contracts/mutation-safety.json"
		if !seen[contractPath] {
			t.Fatal("unsealed contract artifact", contractPath)
		}
		contractBytes, err := os.ReadFile(testutil.ReferencePath(contractPath))
		if err != nil {
			t.Fatal(err)
		}
		contract, err := parseMutationContract(contractBytes)
		if err != nil {
			t.Fatal(err)
		}
		paths := contract.Features
		if contract.Schema == 1 {
			paths = []string{contract.Feature}
		}
		for _, path := range paths {
			if !seen[path] {
				t.Fatal("unsealed contract feature", path)
			}
		}
		if _, err := os.Lstat(testutil.ReferencePath("shared")); !os.IsNotExist(err) {
			t.Fatal("root-only distribution retains shared wrapper", err)
		}
	}
	for _, area := range areas {
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
			if pin.Schema == 1 && rel == "shared/v2/pack/pack-manifest.json" {
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
	if pin.Schema == 1 {
		if _, err := pinnedFile(testutil.ReferencePath("shared", "v2", "pack"), "pack-manifest.json", pin.Pack); err != nil {
			t.Fatal(err)
		}
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
