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

// Both releases use the same native fixture/readback assertions. The current
// candidate compiles Gherkin directly and is sealed by the root manifest.
func TestSharedPackV2(t *testing.T) {
	pin := loadReferencePin(t)
	contract, cases, err := loadMutationContract(pin)
	if err != nil {
		t.Fatal(err)
	}
	fixtureHashes := map[string]string{}
	readbacks := map[string]any{}
	for _, f := range contract.Fixtures {
		t.Run(f.ID, func(t *testing.T) {
			path, err := testutil.LookupFixture(f.AssetID)
			if err != nil {
				t.Fatal(err)
			}
			data, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			fixtureHashes[f.ID] = sha256hex(data)
			members, err := zipPayloads(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(members) != len(f.Members) {
				t.Fatal("member inventory differs")
			}
			for name, hash := range f.Members {
				if sha256hex(members[name]) != hash {
					t.Fatalf("member hash %s", name)
				}
			}
			preserve, err := preservedMembers(f)
			if err != nil {
				t.Fatal(err)
			}
			for name, hash := range preserve {
				if f.Members[name] != hash {
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
			if err := verifySharedFacts(f, readback, members); err != nil {
				t.Fatal(err)
			}
			// A contract edit cannot invent a measured fact even if its JSON is valid.
			bad := f
			bad.Facts = map[string]any{"unproved": true}
			if err := verifySharedFacts(bad, readback, members); err == nil {
				t.Fatal("unproved facts accepted")
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
	dir := os.Getenv("OOXML_REPORT_DIR")
	if dir == "" {
		dir = "../reports/acceptance"
	}
	if err = os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(dir, "shared-v2-verification.json"), map[string]any{"schema": 1, "contractRevision": contract.Revision, "referenceCommit": pin.Commit, "rootManifestSHA256": pin.Manifest, "compiledCases": cases, "fixtures": fixtureHashes, "verifiedFixtureCount": 4, "nativeReadbacks": readbacks, "inventoriedWorkflowCases": 19, "executedWorkflowCases": 0, "status": "fixture-and-contract-integrity-only", "subject": map[string]string{"kind": "native-library", "transport": "none"}})
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
