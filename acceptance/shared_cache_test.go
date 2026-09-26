package acceptance

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

// Direct native execution of the pinned shared cache contract. This is separate
// from Godog inventory and not a schema2 cross-runtime outcome until that adapter
// is implemented. Every observable assertion is checked against actual output.
func TestSharedCacheInvalidation(t *testing.T) {
	root := testutil.ReferencePath("shared", "v2", "pack")
	sharedPackHash := loadReferencePin(t).Pack
	if _, err := pinnedFile(root, "pack-manifest.json", sharedPackHash); err != nil {
		t.Fatal(err)
	}
	const inputHash = "8ba5708d5030adf93a4f7e4ae466563a1b66341067a200a6cb9b782484dcb5b1"
	path, err := testutil.LookupFixture("fixture-" + inputHash)
	if err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(path)
	if err == nil && sha256hex(data) != inputHash {
		t.Fatal("shared cache input hash differs")
	}
	if err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	source := filepath.Join(dir, "source.xlsx")
	output := filepath.Join(dir, "output.xlsx")
	if err = os.WriteFile(source, data, 0600); err != nil {
		t.Fatal(err)
	}
	s, err := spreadsheet.OpenEditing(data, packaging.Limits{MaxSourceBytes: 64 << 20, MaxEntries: 4096, MaxPartBytes: 32 << 20, MaxTotalBytes: 128 << 20})
	if err != nil {
		t.Fatal(err)
	}
	target, err := s.FindNumber("Input", "A1")
	if err != nil {
		t.Fatal(err)
	}
	effect, err := s.SetNumberWithInvalidation(target, 10)
	if err != nil {
		t.Fatal(err)
	}
	receipt, err := s.SaveAs(output)
	if err != nil {
		t.Fatal(err)
	}
	actualSource, err := os.ReadFile(source)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(actualSource, data) {
		t.Fatal("source mutated")
	}
	actual, err := os.ReadFile(output)
	if err != nil {
		t.Fatal(err)
	}
	members, err := zipPayloads(actual)
	if err != nil {
		t.Fatal(err)
	}
	original, err := zipPayloads(data)
	if err != nil {
		t.Fatal(err)
	}
	committed := 0
	input, _, err := sharedCell(members["xl/worksheets/sheet1.xml"], "A1")
	if err != nil {
		t.Fatal(err)
	}
	if effect.ValueChanged && input == "10" {
		committed++
	}
	if committed != 1 {
		t.Fatal("one applied edit not delivered")
	}
	cached, formula, err := sharedCell(members["xl/worksheets/sheet2.xml"], "A1")
	if err != nil {
		t.Fatal(err)
	}
	if formula != "Input!A1*2" || cached != "" {
		t.Fatalf("formula/cache %q/%q", formula, cached)
	}
	if effect.State != "recalculation-required" {
		t.Fatal("missing calculation state")
	}
	d, err := losslessxml.Parse(members["xl/workbook.xml"])
	if err != nil {
		t.Fatal(err)
	}
	flags := map[string]string{}
	for _, e := range d.Elements() {
		if e.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "calcPr"}) {
			for _, a := range e.Attributes() {
				flags[a.Name.Local] = a.Value
			}
		}
	}
	if flags["calcMode"] != "auto" || flags["fullCalcOnLoad"] != "1" || flags["forceFullCalc"] != "1" {
		t.Fatal("recalculation flags missing")
	}
	allowed := map[string]bool{"xl/worksheets/sheet1.xml": true, "xl/worksheets/sheet2.xml": true, "xl/workbook.xml": true}
	if len(original) != len(members) {
		t.Fatal("part set changed")
	}
	changed := []string{}
	for name, before := range original {
		after, exists := members[name]
		if !exists {
			t.Fatal("part removed", name)
		}
		if !bytes.Equal(before, after) {
			if !allowed[name] {
				t.Fatal("unallowed changed part", name)
			}
			changed = append(changed, name)
		}
	}
	if len(changed) != 3 || len(receipt.Changes) != 3 {
		t.Fatal("unexpected changed part inventory", changed, receipt)
	}
	q, err := packaging.OpenPreserved(actual, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	graph, err := q.Graph()
	if err != nil {
		t.Fatal(err)
	}
	sentinel := false
	for _, e := range graph.Edges {
		if e.ResolvedPart == "customXml/preservation-sentinel.xml" {
			sentinel = true
		}
	}
	if !sentinel {
		t.Fatal("sentinel ownership lost")
	}
	reopened, err := spreadsheet.OpenEditing(actual, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	if err = reopened.ValidateStyles(); err != nil {
		t.Fatal(err)
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 2 {
		t.Fatal("staging residue")
	}
	reportDir := os.Getenv("OOXML_REPORT_DIR")
	if reportDir == "" {
		reportDir = "../reports/acceptance"
	}
	if err = os.MkdirAll(reportDir, 0755); err != nil {
		t.Fatal(err)
	}
	artifact := filepath.Join(reportDir, "shared-cache-output.xlsx")
	if err = os.WriteFile(artifact, actual, 0600); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(reportDir, "shared-cache-execution.json"), map[string]any{"schema": 1, "contractRevision": "ooxml-shared-contracts-v2", "scenarioId": "@id-xlsx-cross-sheet-cache-invalidation", "stableCaseKey": "@id-xlsx-cross-sheet-cache-invalidation:{}", "binding": "native Go direct contract test; shared Godog/schema2 adapter pending", "subject": map[string]string{"kind": "native-library", "transport": "none"}, "fixtureSHA256": inputHash, "packManifestSHA256": sharedPackHash, "committedChanges": committed, "sourceUnchanged": true, "input": input, "formula": formula, "cachedValue": cached, "calculationState": effect.State, "calculationPerformed": false, "invalidated": effect.Invalidated, "receipt": receipt, "allReferencesResolve": true, "unrelatedMembersIdentical": true, "outputSHA256": sha256hex(actual)})
	// Ensure persisted report remains ordinary strict JSON rather than NaN/null coercion.
	b, err := os.ReadFile(filepath.Join(reportDir, "shared-cache-execution.json"))
	if err != nil || !json.Valid(b) {
		t.Fatal("invalid report", err)
	}
}
