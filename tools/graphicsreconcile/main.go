// graphicsreconcile matches sealed graphics recipes to actual go test -json
// results. It gives no shared lifecycle or Gherkin-step execution credit.
package main

import (
	"bufio"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

const sharedCommit = "0cf83c156c8e9433347e446e58c415db011b17be"
const manifestSHA = "05d39331f9b832b54716139e57c7f5a324a0a083eab6a12de3d61f4fe74b1967"

var bindings = []struct{ ledger, test string }{
	{"picture-inspection", "TestGraphicsPictureInspectionRecipes"},
	{"picture-insertion", "TestGraphicsPictureInsertionRecipes"},
	{"picture-replacement", "TestGraphicsPictureReplacementRecipes"},
	{"picture-crop", "TestGraphicsPictureMetadataRecipes"},
	{"picture-transform", "TestGraphicsPictureMetadataRecipes"},
	{"picture-placement", "TestGraphicsPicturePlacementRecipes"},
	{"picture-svg", "TestGraphicsPictureSVGRecipes"},
	{"picture-delete", "TestGraphicsPictureDeletionRecipes"},
	{"shape-group", "TestGraphicsShapeGroupingRecipes"},
	{"group-transform", "TestGraphicsGroupTransformRecipes"},
	{"connectors", "TestGraphicsConnectorRecipes"},
	{"diagrams", "TestGraphicsDiagramRecipes"},
	{"smartart-inspection", "TestGraphicsSmartArtInspectionRecipes"},
	{"smartart-copy", "TestGraphicsSmartArtCopyRecipes"},
	{"autoshapes", "TestGraphicsAutoShapeRecipes"},
	{"freeform", "TestGraphicsFreeformRecipes"},
	{"z-order", "TestGraphicsZOrderRecipes"},
	{"gradients", "TestGraphicsGradientRecipes"},
	{"opacity", "TestGraphicsOpacityRecipes"},
	{"outlines", "TestGraphicsOutlineRecipes"},
	{"smartart-office-source", ""},
}

type asset struct {
	ID    string `json:"id"`
	Path  string `json:"path"`
	Bytes int    `json:"bytes"`
	SHA   string `json:"sha256"`
}
type row struct {
	Scenario string `json:"scenarioId"`
	Case     string `json:"caseId"`
	Leaf     string `json:"goLeaf"`
	Status   string `json:"status"`
}
type slice struct {
	Ledger   string `json:"ledger"`
	Feature  string `json:"feature"`
	Contract string `json:"contract"`
	Cases    []row  `json:"cases"`
}
type report struct {
	Schema         int      `json:"schemaVersion"`
	Status         string   `json:"status"`
	SharedCommit   string   `json:"sharedCommit"`
	SharedManifest string   `json:"sharedManifestSHA256"`
	OriginalCases  int      `json:"originalCases"`
	OriginalIDs    int      `json:"originalIDs"`
	SourceCases    int      `json:"sealedSourceCases"`
	SourceIDs      int      `json:"sealedSourceIDs"`
	Assets         []asset  `json:"sealedAssets"`
	Slices         []slice  `json:"slices"`
	Limits         []string `json:"limits"`
}

func goName(s string) string {
	return strings.Map(func(r rune) rune {
		if r == ' ' || r == '\t' || r == '\n' || r == '\r' || r == '\v' || r == '\f' {
			return '_'
		}
		return r
	}, s)
}
func readEvents(path string) (map[string][]string, error) {
	f, e := os.Open(path)
	if e != nil {
		return nil, e
	}
	defer f.Close()
	results := map[string][]string{}
	scanner := bufio.NewScanner(f)
	scanner.Buffer(make([]byte, 65536), 4<<20)
	packages := map[string]string{}
	for scanner.Scan() {
		var event struct{ Action, Package, Test string }
		if e = json.Unmarshal(scanner.Bytes(), &event); e != nil {
			return nil, e
		}
		if event.Action == "fail" || event.Action == "skip" {
			if event.Test == "" {
				return nil, fmt.Errorf("package %s %s", event.Package, event.Action)
			}
		}
		if event.Test == "" && (event.Action == "pass" || event.Action == "fail") {
			packages[event.Package] = event.Action
		}
		if event.Package == "github.com/rcarmo/go-ooxml/pkg/presentation" && (event.Action == "pass" || event.Action == "skip" || event.Action == "fail") {
			results[event.Test] = append(results[event.Test], event.Action)
		}
	}
	if e = scanner.Err(); e != nil {
		return nil, e
	}
	for _, p := range []string{"pkg/presentation", "pkg/packaging", "internal/losslessxml"} {
		if packages["github.com/rcarmo/go-ooxml/"+p] != "pass" {
			return nil, fmt.Errorf("related package %s has no passing completion", p)
		}
	}
	return results, nil
}
func exactlyOnePass(actions []string) bool { return len(actions) == 1 && actions[0] == "pass" }

func reconcile(root, events string) (report, error) {
	out := report{Schema: 1, Status: "passed", SharedCommit: sharedCommit, SharedManifest: manifestSHA, Limits: []string{"Native recipe leaf execution only; no shared lifecycle or Gherkin-step execution credit", "Hashes check sealed input provenance only; generated outputs are checked by semantic and lexical custody tests", "LibreOffice quality results are recorded separately; no Microsoft application certification or general visual equivalence"}}
	raw, e := os.ReadFile(filepath.Join(root, "manifest.json"))
	if e != nil {
		return out, e
	}
	h := sha256.Sum256(raw)
	if hex.EncodeToString(h[:]) != manifestSHA {
		return out, fmt.Errorf("manifest seal differs")
	}
	var manifest struct {
		Files []asset `json:"files"`
	}
	if e = json.Unmarshal(raw, &manifest); e != nil {
		return out, e
	}
	index := map[string]asset{}
	for _, a := range manifest.Files {
		index[a.ID] = a
		index[a.Path] = a
	}
	verified := map[string]asset{}
	read := func(key string) ([]byte, error) {
		a, ok := index[key]
		if !ok || !filepath.IsLocal(a.Path) {
			return nil, fmt.Errorf("unsealed input %s", key)
		}
		info, e := os.Lstat(filepath.Join(root, a.Path))
		if e != nil {
			return nil, e
		}
		if !info.Mode().IsRegular() {
			return nil, fmt.Errorf("nonregular input %s", a.Path)
		}
		b, e := os.ReadFile(filepath.Join(root, a.Path))
		if e != nil {
			return nil, e
		}
		h := sha256.Sum256(b)
		if len(b) != a.Bytes || hex.EncodeToString(h[:]) != a.SHA {
			return nil, fmt.Errorf("asset seal differs %s", a.Path)
		}
		verified[a.Path] = a
		return b, nil
	}
	results, e := readEvents(events)
	if e != nil {
		return out, e
	}
	expected := map[string]bool{}
	originalIDs, sourceIDs := map[string]bool{}, map[string]bool{}
	for _, binding := range bindings {
		path := "ledgers/pptx-" + binding.ledger + ".json"
		raw, e := read(path)
		if e != nil {
			return out, e
		}
		var recipe struct {
			Feature, Contract string
			Cases             []row `json:"cases"`
		}
		if e = json.Unmarshal(raw, &recipe); e != nil {
			return out, e
		}
		feature, e := read(recipe.Feature)
		if e != nil {
			return out, e
		}
		if _, e = read(recipe.Contract); e != nil {
			return out, e
		}
		// Verify every directly referenced sealed asset, including request media.
		var value any
		if e = json.Unmarshal(raw, &value); e != nil {
			return out, e
		}
		var visit func(any) error
		visit = func(v any) error {
			switch x := v.(type) {
			case string:
				if _, ok := index[x]; ok {
					_, e := read(x)
					return e
				}
				if strings.HasPrefix(x, "fixture-") || strings.HasPrefix(x, "asset-") {
					return fmt.Errorf("unknown asset %s", x)
				}
			case []any:
				for _, v := range x {
					if e := visit(v); e != nil {
						return e
					}
				}
			case map[string]any:
				for _, v := range x {
					if e := visit(v); e != nil {
						return e
					}
				}
			}
			return nil
		}
		if e = visit(value); e != nil {
			return out, e
		}
		s := slice{Ledger: path, Feature: recipe.Feature, Contract: recipe.Contract}
		for _, c := range recipe.Cases {
			if c.Scenario == "" || c.Case == "" || !strings.Contains(string(feature), c.Scenario) {
				return out, fmt.Errorf("recipe missing feature identity %s", path)
			}
			test := binding.test
			if test == "" {
				if c.Scenario == "@id-pptx-smartart-office-source-inspection" {
					test = "TestGraphicsSmartArtSealedSourceInspection"
				} else if c.Scenario == "@id-pptx-smartart-office-source-copy" {
					test = "TestGraphicsSmartArtSealedSourceCopies"
				} else {
					return out, fmt.Errorf("unbound source scenario")
				}
				out.SourceCases++
				sourceIDs[c.Scenario] = true
			} else {
				out.OriginalCases++
				originalIDs[c.Scenario] = true
			}
			c.Leaf = test + "/" + goName(c.Scenario+" ["+c.Case+"]")
			if expected[c.Leaf] {
				return out, fmt.Errorf("duplicate binding %s", c.Leaf)
			}
			expected[c.Leaf] = true
			got := results[c.Leaf]
			if !exactlyOnePass(got) {
				return out, fmt.Errorf("leaf %s needs exactly one pass; got %v", c.Leaf, got)
			}
			c.Status = "passed"
			s.Cases = append(s.Cases, c)
		}
		out.Slices = append(out.Slices, s)
	}
	// Reject accidental unmatched recipe leaves rather than counting parent passes.
	for test := range results {
		if strings.Contains(test, "/@id-") && !expected[test] {
			return out, fmt.Errorf("unmatched recipe leaf %s", test)
		}
	}
	out.OriginalIDs = len(originalIDs)
	out.SourceIDs = len(sourceIDs)
	if out.OriginalCases != 320 || out.OriginalIDs != 54 || out.SourceCases != 3 || out.SourceIDs != 2 {
		return out, fmt.Errorf("graphics inventory differs: %d/%d + %d/%d", out.OriginalCases, out.OriginalIDs, out.SourceCases, out.SourceIDs)
	}
	for _, a := range verified {
		out.Assets = append(out.Assets, a)
	}
	sort.Slice(out.Assets, func(i, j int) bool { return out.Assets[i].Path < out.Assets[j].Path })
	return out, nil
}
func main() {
	root := flag.String("root", "", "sealed shared checkout")
	events := flag.String("events", "", "go test -json events")
	output := flag.String("output", "", "report path")
	flag.Parse()
	if *root == "" || *events == "" || *output == "" {
		fmt.Fprintln(os.Stderr, "root, events and output required")
		os.Exit(2)
	}
	r, e := reconcile(*root, *events)
	if e != nil {
		fmt.Fprintln(os.Stderr, e)
		os.Exit(1)
	}
	raw, e := json.MarshalIndent(r, "", "  ")
	if e == nil {
		e = os.MkdirAll(filepath.Dir(*output), 0755)
	}
	if e == nil {
		e = os.WriteFile(*output, append(raw, '\n'), 0644)
	}
	if e != nil {
		fmt.Fprintln(os.Stderr, e)
		os.Exit(1)
	}
	fmt.Printf("Reconciled %d/%d original cases/IDs + %d/%d sealed-source cases/IDs; %d sealed assets\n", r.OriginalCases, r.OriginalIDs, r.SourceCases, r.SourceIDs, len(r.Assets))
}
