package acceptance

import (
	"encoding/json"
	"fmt"
	"go/ast"
	"go/constant"
	"go/parser"
	"go/token"
	"go/types"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// A sealed workflow may retain historical prose evidence or the newer list
// form (empty for planned cases). Reject all other JSON types explicitly.
func workflowEvidencePresent(raw json.RawMessage) (bool, error) {
	if len(raw) == 0 || strings.TrimSpace(string(raw)) == "null" {
		return false, fmt.Errorf("workflow evidence missing or null")
	}
	var prose string
	if err := json.Unmarshal(raw, &prose); err == nil {
		return len(strings.TrimSpace(prose)) > 0, nil
	}
	var entries []string
	if err := json.Unmarshal(raw, &entries); err != nil || entries == nil {
		return false, fmt.Errorf("workflow evidence must be string or string array")
	}
	for _, e := range entries {
		if strings.TrimSpace(e) == "" {
			return false, fmt.Errorf("workflow evidence has empty item")
		}
	}
	return len(entries) > 0, nil
}

func TestWorkflowEvidenceTypes(t *testing.T) {
	for _, tc := range []struct {
		input            string
		present, invalid bool
	}{
		{`"historical source"`, true, false}, {`[]`, false, false}, {`["source","receipt"]`, true, false},
		{`" "`, false, false}, {`[""]`, false, true}, {`null`, false, true}, {`{}`, false, true}, {`42`, false, true}, {`[3]`, false, true}, {``, false, true},
	} {
		got, err := workflowEvidencePresent(json.RawMessage(tc.input))
		if (err != nil) != tc.invalid || (!tc.invalid && got != tc.present) {
			t.Errorf("evidence %q present=%v error=%v", tc.input, got, err)
		}
	}
	present, err := workflowEvidencePresent(json.RawMessage(`[]`))
	if err != nil || present {
		t.Fatalf("planned empty evidence: %v %v", present, err)
	}
	present, err = workflowEvidencePresent(json.RawMessage(`" "`))
	if err != nil || present {
		t.Fatalf("implemented empty evidence accepted: %v %v", present, err)
	}
}

type referenceFact struct {
	ID       string   `json:"id"`
	Value    string   `json:"value"`
	Status   string   `json:"status"`
	Evidence []string `json:"evidence"`
}

func readReferenceJSON(name string, target any) error {
	if !filepath.IsLocal(name) {
		return fmt.Errorf("non-local reference path %s", name)
	}
	b, err := os.ReadFile(testutil.ReferencePath(name))
	if err != nil {
		return err
	}
	return json.Unmarshal(b, target)
}

// Registry consistency is independent of execution or specification authority.
// Existing Go constants remain observed declarations, including disputed aliases.
func TestSharedFactRegistry(t *testing.T) {
	var evidence struct {
		Schema   int `json:"schemaVersion"`
		Evidence []struct {
			ID   string `json:"id"`
			Kind string `json:"kind"`
		} `json:"items"`
	}
	if err := readReferenceJSON("facts/evidence.json", &evidence); err != nil {
		t.Fatal(err)
	}
	if evidence.Schema != 1 || len(evidence.Evidence) == 0 {
		t.Fatal("missing evidence registry")
	}
	ids := map[string]bool{}
	for _, e := range evidence.Evidence {
		if e.ID == "" || e.Kind == "" || ids[e.ID] {
			t.Fatalf("invalid/duplicate evidence %+v", e)
		}
		ids[e.ID] = true
	}
	facts := map[string]referenceFact{}
	for _, path := range []string{"content-types", "namespaces", "relationships", "constants"} {
		var registry struct {
			Schema int             `json:"schemaVersion"`
			Values []referenceFact `json:"values"`
		}
		if err := readReferenceJSON("facts/"+path+".json", &registry); err != nil {
			t.Fatal(err)
		}
		if registry.Schema != 1 || len(registry.Values) == 0 {
			t.Fatal("empty fact registry", path)
		}
		for _, f := range registry.Values {
			if f.ID == "" || f.Value == "" || facts[f.ID].ID != "" || len(f.Evidence) == 0 {
				t.Fatalf("invalid/duplicate fact %+v", f)
			}
			switch f.Status {
			case "observed", "specified", "disputed":
			default:
				t.Fatalf("unknown fact status %+v", f)
			}
			for _, id := range f.Evidence {
				if !ids[id] {
					t.Fatalf("dangling fact evidence %s -> %s", f.ID, id)
				}
			}
			facts[f.ID] = f
		}
	}
	// Type-check the single native declaration file, including constant aliases.
	// There are no imports/build conditions in that file; no runtime dependency on
	// source parsers or the shared registry is introduced by this test.
	fs := token.NewFileSet()
	file, err := parser.ParseFile(fs, "../pkg/packaging/constants.go", nil, 0)
	if err != nil {
		t.Fatal(err)
	}
	conf := types.Config{}
	pkg, err := conf.Check("packaging", fs, []*ast.File{file}, nil)
	if err != nil {
		t.Fatal(err)
	}
	aliases := map[string]string{
		"ContentTypeWordFooter": "ContentTypeFooter",
		"ContentTypeWordHeader": "ContentTypeHeader",
		"ContentTypeWordStyles": "ContentTypeStyles",
		"PPTXPresentationPath":  "PresentationPath",
	}
	compared := 0
	for _, name := range pkg.Scope().Names() {
		c, ok := pkg.Scope().Lookup(name).(*types.Const)
		if !ok || !c.Exported() || c.Val().Kind() != constant.String {
			continue
		}
		expected := constant.StringVal(c.Val())
		id := name
		if canonical, alias := aliases[name]; alias {
			id = canonical
		}
		f, ok := facts[id]
		if !ok || f.Value != expected {
			t.Errorf("native fact %s missing or differs: Go=%q registry=%q", name, expected, f.Value)
		}
		compared++
	}
	if compared == 0 {
		t.Fatal("no native constants checked")
	}
	old, specified := facts["ContentTypeCommentsExtended"], facts["ContentTypeCommentsExtendedSpecified"]
	if old.Status != "disputed" || specified.Status != "specified" || old.Value == specified.Value {
		t.Fatal("disputed and specified MIME facts collapsed")
	}
	t.Logf("%d native string constants checked against %d evidence-linked shared facts; no execution credit", compared, len(facts))
}

func TestSharedWorkflowRegistry(t *testing.T) {
	var ledger struct {
		Schema    int      `json:"schemaVersion"`
		Features  []string `json:"features"`
		Workflows []struct {
			ID        string   `json:"id"`
			Feature   string   `json:"feature"`
			Cases     int      `json:"expandedCases"`
			Facts     []string `json:"factIds"`
			Expected  []string `json:"expectedOutcomes"`
			Consumers map[string]struct {
				State    string          `json:"status"`
				Evidence json.RawMessage `json:"evidence"`
			} `json:"consumers"`
		} `json:"workflows"`
	}
	if err := readReferenceJSON("ledgers/workflows.json", &ledger); err != nil {
		t.Fatal(err)
	}
	if ledger.Schema != 1 || len(ledger.Workflows) == 0 {
		t.Fatal("missing workflow registry")
	}
	features := map[string]bool{}
	for _, name := range ledger.Features {
		if !filepath.IsLocal(name) || features[name] {
			t.Fatal("invalid feature path", name)
		}
		if _, err := os.ReadFile(testutil.ReferencePath(name)); err != nil {
			t.Fatal(err)
		}
		features[name] = true
	}
	facts := map[string]bool{}
	for _, path := range []string{"content-types", "namespaces", "relationships", "constants"} {
		var r struct {
			Values []referenceFact `json:"values"`
		}
		if err := readReferenceJSON("facts/"+path+".json", &r); err != nil {
			t.Fatal(err)
		}
		for _, f := range r.Values {
			facts[f.ID] = true
		}
	}
	seen := map[string]bool{}
	cases := 0
	for _, w := range ledger.Workflows {
		if w.ID == "" || seen[w.ID] || !features[w.Feature] || w.Cases < 1 || len(w.Expected) == 0 {
			t.Fatalf("invalid workflow %+v", w)
		}
		seen[w.ID] = true
		cases += w.Cases
		for _, id := range w.Facts {
			if !facts[id] {
				t.Fatalf("unknown workflow fact %s -> %s", w.ID, id)
			}
		}
		goState, ok := w.Consumers["go"]
		if !ok {
			t.Fatalf("missing Go accounting %s", w.ID)
		}
		hasEvidence, evidenceErr := workflowEvidencePresent(goState.Evidence)
		if evidenceErr != nil {
			t.Fatalf("invalid Go evidence %s: %v", w.ID, evidenceErr)
		}
		switch goState.State {
		case "planned", "unmapped": // Explicitly no local execution credit.
		case "implemented":
			if !hasEvidence {
				t.Fatalf("missing claimed Go evidence %s", w.ID)
			}
		default:
			t.Fatalf("unknown Go accounting %s: %s", w.ID, goState.State)
		}
	}
	t.Logf("%d workflow IDs/%d expanded cases inventoried; no workflow execution credit", len(seen), cases)
}
