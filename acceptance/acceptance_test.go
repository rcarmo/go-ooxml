package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// A case is identified by file, stable scenario tag and source line of its
// Examples row (or scenario). Counts alone cannot establish execution coverage.
type caseID struct {
	File string `json:"file"`
	ID   string `json:"id"`
	Line int    `json:"line"`
}
type expectedCase struct {
	caseID
	Name  string `json:"name"`
	Steps int    `json:"steps"`
}
type reportStep struct {
	Name   string `json:"name"`
	Result struct {
		Status string `json:"status"`
	} `json:"result"`
}
type reportElement struct {
	Name string `json:"name"`
	Line int    `json:"line"`
	Type string `json:"type"`
	Tags []struct {
		Name string `json:"name"`
	} `json:"tags"`
	Steps []reportStep `json:"steps"`
}
type reportFeature struct {
	URI      string          `json:"uri"`
	Elements []reportElement `json:"elements"`
}

type world struct {
	source   []byte
	pkg      *packaging.Package
	fixtures map[string]string
}

func (w *world) fixture(name string) error {
	if !filepath.IsLocal(name) {
		return fmt.Errorf("unsafe fixture %q", name)
	}
	data, err := os.ReadFile(testutil.FixturePath(name))
	if err != nil {
		return err
	}
	w.source = data
	sum := sha256.Sum256(data)
	w.fixtures[name] = hex.EncodeToString(sum[:])
	return nil
}
func (w *world) open() error { var err error; w.pkg, err = packaging.OpenBytes(w.source); return err }
func (w *world) contains(name string) error {
	if w.pkg == nil {
		return fmt.Errorf("package not open")
	}
	p, err := w.pkg.GetPart(name)
	if err != nil {
		return err
	}
	content, err := p.Content()
	if err != nil {
		return err
	}
	if len(content) == 0 {
		return fmt.Errorf("empty part %s", name)
	}
	return nil
}

func TestAcceptance(t *testing.T) {
	w := &world{fixtures: map[string]string{}}
	initializer := func(sc *godog.ScenarioContext) {
		safetySteps(sc)
		zip64Steps(sc)
		overlapSteps(sc)
		limitSteps(sc)
		preservedSteps(sc)
		receiptSteps(sc)
		xmlSteps(sc)
		graphSteps(sc)
		graphEditSteps(sc)
		graphDeleteSteps(sc)
		graphTargetFormSteps(sc)
		wordSteps(sc)
		slideSteps(sc)
		notesEditSteps(sc)
		notesMultilineSteps(sc)
		cellSteps(sc)
		vocabularySteps(sc)
		storySteps(sc)
		spanSteps(sc)
		replaceSpanSteps(sc)
		spanBatchSteps(sc)
		searchPolicySteps(sc)
		attributeSteps(sc)
		whitespaceSteps(sc)
		replaceAllSteps(sc)
		searchScopeSteps(sc)
		runBoundarySteps(sc)
		spanGuardSteps(sc)
		styleIndexSteps(sc)
		xmlNormalizationSteps(sc)
		xmlInsertSteps(sc)
		subtreeSteps(sc)
		xmlConformanceSteps(sc)
		officeLinkSteps(sc)
		formulaSteps(sc)
		cacheSteps(sc)
		imageReplaceSteps(sc)
		remapSteps(sc)
		commentMIMESteps(sc)
		sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
			w.source = nil
			w.pkg = nil
			return ctx, nil
		})
		sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
			if w.pkg != nil {
				_ = w.pkg.Close()
			}
			return ctx, nil
		})
		sc.Step(`^the repository fixture "([^"]+)"$`, w.fixture)
		sc.Step(`^I open its Office package$`, w.open)
		sc.Step(`^the package contains a nonempty part "([^"]+)"$`, w.contains)
	}
	expected, inventory, err := inventoryCases()
	if err != nil {
		t.Fatal(err)
	}
	if len(expected) == 0 {
		t.Fatal("no implemented cases inventoried")
	}
	var combined []reportFeature
	for _, selection := range []struct {
		name, path, tags string
	}{
		{"go-ooxml-native", goFeatureRoot(), "@implemented && @go"},
		{"go-ooxml-overlap", overlapFeaturePath(), overlapCaseID},
	} {
		var output bytes.Buffer
		suite := godog.TestSuite{Name: selection.name, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{selection.path}, Tags: selection.tags, Strict: true, Concurrency: 1}, ScenarioInitializer: initializer}
		code := suite.Run()
		var features []reportFeature
		if err := json.Unmarshal(output.Bytes(), &features); err != nil {
			t.Fatalf("%s invalid report: %v (%s)", selection.name, err, output.String())
		}
		if code != 0 {
			t.Errorf("%s Godog failed (%d): %s", selection.name, code, output.String())
		}
		combined = append(combined, features...)
	}
	data, err := json.Marshal(combined)
	if err != nil {
		t.Fatal(err)
	}
	dir := os.Getenv("OOXML_REPORT_DIR")
	if dir == "" {
		dir = "../reports/acceptance"
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "cucumber.json"), data, 0644); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(dir, "inventory.json"), inventory)
	writeJSON(t, filepath.Join(dir, "environment.json"), map[string]any{"go": runtime.Version(), "fixture_sha256": w.fixtures, "scope": "implemented native plus one exact canonical overlap case", "external_executed": false})
	if err := reconcile(expected, data); err != nil {
		t.Error(err)
	}
}

func writeJSON(t *testing.T, path string, value any) {
	t.Helper()
	b, err := json.MarshalIndent(value, "", "  ")
	if err != nil {
		t.Fatal(err)
	}
	if err = os.WriteFile(path, append(b, '\n'), 0644); err != nil {
		t.Fatal(err)
	}
}
func stableID(tags []string) (string, error) {
	id := ""
	for _, tag := range tags {
		if nativeIDPattern.MatchString(tag) || tag == overlapCaseID {
			if id != "" {
				return "", fmt.Errorf("multiple IDs: %v", tags)
			}
			id = tag
		}
	}
	if id == "" {
		return "", fmt.Errorf("missing stable ID: %v", tags)
	}
	return id, nil
}

func reconcile(expected map[caseID]expectedCase, data []byte) error {
	var features []reportFeature
	if err := json.Unmarshal(data, &features); err != nil {
		return err
	}
	seen := map[caseID]bool{}
	for _, f := range features {
		for _, e := range f.Elements {
			if e.Type != "scenario" {
				return fmt.Errorf("unexpected report element %q", e.Type)
			}
			tags := []string{}
			for _, tag := range e.Tags {
				tags = append(tags, tag.Name)
			}
			id, err := stableID(tags)
			if err != nil {
				return err
			}
			key := caseID{filepath.ToSlash(filepath.Clean(f.URI)), id, e.Line}
			want, ok := expected[key]
			if !ok {
				return fmt.Errorf("unplanned result %+v", key)
			}
			if seen[key] {
				return fmt.Errorf("duplicate result %+v", key)
			}
			seen[key] = true
			if e.Name != want.Name || len(e.Steps) != want.Steps {
				return fmt.Errorf("name/steps mismatch %+v", key)
			}
			for _, step := range e.Steps {
				if step.Result.Status != "passed" {
					return fmt.Errorf("%+v step %q: %s", key, step.Name, step.Result.Status)
				}
			}
		}
	}
	for key := range expected {
		if !seen[key] {
			return fmt.Errorf("missing result %+v", key)
		}
	}
	return nil
}
