package acceptance

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"sort"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

func sha256hex(data []byte) string { s := sha256.Sum256(data); return hex.EncodeToString(s[:]) }
func pinnedFile(root, name, hash string) ([]byte, error) {
	if !filepath.IsLocal(name) {
		return nil, fmt.Errorf("non-local pack member %s", name)
	}
	b, err := os.ReadFile(filepath.Join(root, name))
	if err != nil {
		return nil, err
	}
	if sha256hex(b) != hash {
		return nil, fmt.Errorf("pack hash mismatch %s", name)
	}
	return b, nil
}
func caseKey(id string, values map[string]string) string {
	keys := []string{}
	for key := range values {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	var b bytes.Buffer
	b.WriteByte('{')
	for i, key := range keys {
		if i > 0 {
			b.WriteByte(',')
		}
		k, _ := json.Marshal(key)
		v, _ := json.Marshal(values[key])
		b.Write(k)
		b.WriteByte(':')
		b.Write(v)
	}
	b.WriteByte('}')
	return id + ":" + b.String()
}

type sharedFixture struct {
	ID           string            `json:"id"`
	Path         string            `json:"path"`
	AssetID      string            `json:"assetId"`
	SHA256       string            `json:"sha256"`
	MemberSHA256 map[string]string `json:"memberSha256"`
	Preserve     map[string]string `json:"mustPreservePayloads"`
}

// The shared code-free distribution is required for native fixture/contract
// checks. Integrity verification does not count as workflow execution.
func TestSharedPackV2(t *testing.T) {
	root := testutil.ReferencePath("shared", "v2", "pack")
	sharedPackHash := loadReferencePin(t).Pack
	b, err := pinnedFile(root, "pack-manifest.json", sharedPackHash)
	if err != nil {
		t.Fatal(err)
	}
	var pack struct {
		Revision      string            `json:"contractRevision"`
		Distribution  string            `json:"distributionRevision"`
		ScenarioCount int               `json:"scenarioCount"`
		CaseCount     int               `json:"expandedCaseCount"`
		Files         map[string]string `json:"files"`
	}
	if err = json.Unmarshal(b, &pack); err != nil {
		t.Fatal(err)
	}
	if pack.Revision != "ooxml-shared-contracts-v2" || pack.Distribution != "fixtures-ooxml-v0.2.0" || pack.ScenarioCount != 8 || pack.CaseCount != 19 {
		t.Fatal("wrong shared revision/inventory")
	}
	for name, hash := range pack.Files {
		if _, err = pinnedFile(root, name, hash); err != nil {
			t.Fatal(err)
		}
	}
	b, err = pinnedFile(root, "fixture-manifest.json", pack.Files["fixture-manifest.json"])
	if err != nil {
		t.Fatal(err)
	}
	var fixtures struct {
		Schema   int             `json:"schemaVersion"`
		PathBase string          `json:"pathBase"`
		Fixtures []sharedFixture `json:"fixtures"`
	}
	if err = json.Unmarshal(b, &fixtures); err != nil {
		t.Fatal(err)
	}
	if fixtures.Schema != 2 || fixtures.PathBase != "repository-root" || len(fixtures.Fixtures) != 4 {
		t.Fatal("wrong fixture inventory")
	}
	fixtureHashes := map[string]string{}
	readbacks := map[string]any{}
	for _, f := range fixtures.Fixtures {
		t.Run(f.ID, func(t *testing.T) {
			path, err := testutil.LookupFixture(f.AssetID)
			if err != nil {
				t.Fatal(err)
			}
			if path != testutil.ReferencePath(f.Path) {
				t.Fatal("shared path and fixture ID differ")
			}
			data, err := pinnedFile(testutil.ReferenceRoot(), f.Path, f.SHA256)
			if err != nil {
				t.Fatal(err)
			}
			fixtureHashes[f.ID] = f.SHA256
			members, err := zipPayloads(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(members) != len(f.MemberSHA256) {
				t.Fatal("member inventory differs")
			}
			for name, hash := range f.MemberSHA256 {
				if sha256hex(members[name]) != hash {
					t.Fatalf("member hash %s", name)
				}
			}
			for name, hash := range f.Preserve {
				if f.MemberSHA256[name] != hash {
					t.Fatalf("preservation hash inconsistent %s", name)
				}
			}
			p, err := packaging.OpenPreserved(data, packaging.Limits{MaxSourceBytes: 64 << 20, MaxEntries: 4096, MaxPartBytes: 32 << 20, MaxTotalBytes: 128 << 20})
			if err != nil {
				t.Fatal(err)
			}
			g, err := p.Graph()
			if err != nil {
				t.Fatal(err)
			}
			linked := false
			for _, e := range g.Edges {
				if !e.External && e.ResolvedPart == "customXml/preservation-sentinel.xml" {
					linked = true
				}
			}
			if !linked {
				t.Fatal("sentinel relationship absent")
			}
			var out bytes.Buffer
			if err = p.WriteTo(&out); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(data, out.Bytes()) {
				t.Fatal("no-op changed fixture")
			}
			readback, err := sharedReadback(f.ID, data, members)
			if err != nil {
				t.Fatal(err)
			}
			readbacks[f.ID] = readback
			if filepath.Ext(f.ID) == ".xlsx" {
				s, err := spreadsheet.OpenEditing(data, packaging.Limits{})
				if err != nil {
					t.Fatal(err)
				}
				if err = s.ValidateStyles(); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
	b, err = pinnedFile(root, "expanded-contracts.json", pack.Files["expanded-contracts.json"])
	if err != nil {
		t.Fatal(err)
	}
	var compiled struct {
		Cases []struct {
			ID            string            `json:"stableCaseKey"`
			ScenarioID    string            `json:"scenarioId"`
			ExampleValues map[string]string `json:"examples"`
			Lifecycle     string            `json:"lifecycle"`
			Execution     string            `json:"execution"`
			Steps         []struct {
				Argument json.RawMessage `json:"argument"`
			} `json:"expandedSteps"`
		} `json:"inventory"`
	}
	if err = json.Unmarshal(b, &compiled); err != nil {
		t.Fatal(err)
	}
	if len(compiled.Cases) != 19 {
		t.Fatalf("compiled cases: %d", len(compiled.Cases))
	}
	seen := map[string]bool{}
	for _, c := range compiled.Cases {
		key := caseKey(c.ScenarioID, c.ExampleValues)
		if c.ID != key || seen[key] {
			t.Fatalf("unstable/duplicate case key %s", key)
		}
		seen[key] = true
		if c.Lifecycle != "planned" || c.Execution != "not-run" {
			t.Fatal("pack overclaims execution")
		}
		for _, step := range c.Steps {
			if err := validateTypedTable(step.Argument); err != nil {
				t.Fatal(err)
			}
		}
	}
	dir := os.Getenv("OOXML_REPORT_DIR")
	if dir == "" {
		dir = "../reports/acceptance"
	}
	if err = os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(dir, "shared-v2-verification.json"), map[string]any{"schema": 1, "contractRevision": pack.Revision, "distributionRevision": pack.Distribution, "packManifestSHA256": sharedPackHash, "fixtures": fixtureHashes, "verifiedFixtureCount": 4, "nativeReadbacks": readbacks, "inventoriedWorkflowCases": 19, "executedWorkflowCases": 0, "status": "fixture-and-contract-integrity-only", "subject": map[string]string{"kind": "native-library", "transport": "none"}})
}
func validateTypedTable(argument json.RawMessage) error {
	if len(argument) == 0 || string(argument) == "null" {
		return nil
	}
	var a struct {
		DataTable [][]string `json:"dataTable"`
	}
	if err := json.Unmarshal(argument, &a); err != nil {
		return err
	}
	if len(a.DataTable) == 0 {
		return nil
	}
	index := -1
	for i, c := range a.DataTable[0] {
		if c == "value_json" {
			index = i
		}
	}
	if index < 0 {
		return nil
	}
	for _, row := range a.DataTable[1:] {
		if index >= len(row) || !json.Valid([]byte(row[index])) {
			return fmt.Errorf("invalid strict typed JSON table cell")
		}
	}
	return nil
}
func TestSharedIdentityAndTypedCells(t *testing.T) {
	if got := caseKey("@id", map[string]string{"z": "last", "a": "first"}); got != `@id:{"a":"first","z":"last"}` {
		t.Fatal(got)
	}
	for _, tc := range []struct {
		value string
		bad   bool
	}{{`"alpha\nbeta"`, false}, {"\"alpha\nbeta\"", true}, {`125`, false}, {`NaN`, true}} {
		arg, _ := json.Marshal(map[string]any{"dataTable": [][]string{{"value_json"}, {tc.value}}})
		if err := validateTypedTable(arg); (err != nil) != tc.bad {
			t.Fatalf("%q err=%v", tc.value, err)
		}
	}
}
