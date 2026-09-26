package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"strings"
	"testing"

	"github.com/cucumber/gherkin/go/v26"
)

// Validate reviewed staging candidates without executing or awarding canonical
// coverage. The shared registry owns the reconciled behaviour specification.
func TestBehaviourCatalogueCandidates(t *testing.T) {
	for _, name := range []string{"formula", "xml", "archive", "package-model", "delivery", "graph", "test-custody", "utilities", "ooxml-model"} {
		t.Run(name, func(t *testing.T) { validateCatalogueFamily(t, name) })
	}
}

func validateCatalogueFamily(t *testing.T, name string) {
	t.Helper()
	const base = "../docs/behaviors/"
	var mapping struct {
		Status    string   `json:"status"`
		Feature   string   `json:"feature"`
		Prefix    string   `json:"native_prefix"`
		Files     []string `json:"native_files"`
		Canonical bool     `json:"canonical_ids_assigned"`
		Execution bool     `json:"execution_credit"`
		Mappings  []struct {
			Native   string   `json:"native_id"`
			Hash     string   `json:"source_sha256"`
			IDs      []string `json:"candidate_ids"`
			Coverage string   `json:"coverage"`
		} `json:"mappings"`
	}
	b, err := os.ReadFile(base + name + "-mapping.json")
	if err != nil {
		t.Fatal(err)
	}
	if err = json.Unmarshal(b, &mapping); err != nil {
		t.Fatal(err)
	}
	if mapping.Status != "central-reconciliation-staging" || mapping.Canonical || mapping.Execution || mapping.Prefix == "" {
		t.Fatal("staging claims canonical/executed status")
	}
	var inv struct {
		Files []struct {
			Path string `json:"path"`
			Hash string `json:"sha256"`
		} `json:"files"`
		Declarations []struct {
			ID   string `json:"id"`
			File string `json:"file"`
		} `json:"declarations"`
	}
	b, err = os.ReadFile(base + "native-inventory.json")
	if err != nil {
		t.Fatal(err)
	}
	if err = json.Unmarshal(b, &inv); err != nil {
		t.Fatal(err)
	}
	hashes := map[string]string{}
	for _, f := range inv.Files {
		hashes[f.Path] = f.Hash
	}
	decls := map[string]string{}
	required := map[string]bool{}
	selected := map[string]bool{}
	for _, file := range mapping.Files {
		selected[file] = true
	}
	for _, d := range inv.Declarations {
		decls[d.ID] = d.File
		if (len(selected) == 0 && strings.HasPrefix(d.File, mapping.Prefix)) || selected[d.File] {
			required[d.ID] = true
		}
	}
	f, err := os.Open(base + mapping.Feature)
	if err != nil {
		t.Fatal(err)
	}
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	_ = f.Close()
	if err != nil {
		t.Fatal(err)
	}
	ids := map[string]bool{}
	for _, p := range gherkin.Pickles(*doc, mapping.Feature, next) {
		id := ""
		for _, tag := range p.Tags {
			if strings.HasPrefix(tag.Name, "@candidate-") {
				if id != "" {
					t.Fatal("multiple candidate IDs")
				}
				id = tag.Name
			}
			if tag.Name == "@implemented" {
				t.Fatal("staging candidate marked implemented")
			}
		}
		if id == "" || ids[id] || len(p.Steps) < 3 {
			t.Fatal("invalid candidate", p.Name)
		}
		ids[id] = true
		for _, s := range p.Steps {
			if strings.Contains(strings.ToLower(s.Text), "test passes") {
				t.Fatal("vacuous outcome", s.Text)
			}
		}
	}
	seen := map[string]bool{}
	mapped := map[string]bool{}
	for _, m := range mapping.Mappings {
		path, ok := decls[m.Native]
		if !ok || seen[m.Native] || hashes[path] != m.Hash || len(m.IDs) == 0 || m.Coverage != "reviewed-family-and-parameter-groups" {
			t.Fatalf("invalid native mapping %+v", m)
		}
		seen[m.Native] = true
		for _, id := range m.IDs {
			if !ids[id] {
				t.Fatal("unknown candidate", id)
			}
			mapped[id] = true
		}
	}
	for id := range required {
		if !seen[id] {
			t.Fatal("missing native declaration", id)
		}
	}
	for id := range ids {
		if !mapped[id] {
			t.Fatal("unmapped candidate", id)
		}
	}
}
