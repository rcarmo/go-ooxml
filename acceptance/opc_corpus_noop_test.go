package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const opcCorpusNoopCaseID = "@id-opc-package-corpus-noop"
const goOrigin = "https://github.com/rcarmo/go-ooxml"
const pythonOrigin = "https://github.com/rcarmo/python-office-mcp-server"
const pythonTestdataAlias = "fixtures/python-office-mcp-server/tests/_templates/testdata/"

var opcCorpusNoopText = []string{
	"the go-ooxml and python-office-mcp-server fixture corpora are enumerated",
	"each OOXML fixture package is opened and serialized without edits through the OPC layer",
	"every reopened package matches its original whole-archive bytes",
	"both fixture corpora contribute their exact known nonzero fixture counts",
}

func guardOPCCorpusNoopCase(id string, p *messages.Pickle, line int) error {
	if id != opcCorpusNoopCaseID || p == nil || line != 8 || p.Name != "Opening and reopening the fixture corpora without edits preserves whole archives" || len(p.AstNodeIds) != 1 || len(p.Steps) != 4 || len(p.Tags) != 2 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != opcCorpusNoopCaseID {
		return fmt.Errorf("OPC corpus-noop identity drift")
	}
	for i, want := range opcCorpusNoopText {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("OPC corpus-noop step %d drift", i+1)
		}
	}
	return nil
}

func guardOPCCorpusNoopRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "OPC package custody, transactions and save destinations" || len(doc.Feature.Tags) != 1 || doc.Feature.Tags[0].Name != "@planned" {
		return fmt.Errorf("OPC corpus-noop feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "OPC package custody and transactional part edits" {
			continue
		}
		if len(child.Rule.Tags) != 0 {
			return fmt.Errorf("OPC corpus-noop rule tags drift")
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("OPC corpus-noop background drift")
			}
			if s := member.Scenario; s != nil {
				for _, tag := range s.Tags {
					if tag.Name == opcCorpusNoopCaseID {
						found++
						if len(s.Tags) != 1 || tag.Location.Line != 7 || s.Location.Line != 8 || len(s.Examples) != 0 || len(s.Steps) != 4 {
							return fmt.Errorf("OPC corpus-noop structure drift")
						}
					}
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("OPC corpus-noop scenario count %d", found)
	}
	return nil
}

func TestOPCCorpusNoopGuard(t *testing.T) {
	path := packagePreservationFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil {
		t.Fatal(err)
	}
	if err := guardOPCCorpusNoopRule(doc); err != nil {
		t.Fatal(err)
	}
	var picked *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == opcCorpusNoopCaseID {
				if picked != nil {
					t.Fatal("duplicate corpus-noop pickle")
				}
				picked = p
			}
		}
	}
	if err := guardOPCCorpusNoopCase(opcCorpusNoopCaseID, picked, 8); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name   string
		change func(*messages.Pickle)
	}{
		{"id", func(p *messages.Pickle) { p.Tags[1].Name = "@id-other" }},
		{"name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "row") }},
		{"argument", func(p *messages.Pickle) { p.Steps[1].Argument = &messages.PickleStepArgument{} }},
		{"tags", func(p *messages.Pickle) { p.Tags = append(p.Tags, &messages.PickleTag{Name: "@extra"}) }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := *picked
			p.AstNodeIds = append([]string(nil), picked.AstNodeIds...)
			p.Tags = append([]*messages.PickleTag(nil), picked.Tags...)
			for i, tag := range p.Tags {
				copyTag := *tag
				p.Tags[i] = &copyTag
			}
			p.Steps = append([]*messages.PickleStep(nil), picked.Steps...)
			for i, step := range p.Steps {
				copyStep := *step
				p.Steps[i] = &copyStep
			}
			tc.change(&p)
			if guardOPCCorpusNoopCase(opcCorpusNoopCaseID, &p, 8) == nil {
				t.Fatal("accepted drift")
			}
		})
	}
	for i := range opcCorpusNoopText {
		p := *picked
		p.Steps = append([]*messages.PickleStep(nil), picked.Steps...)
		s := *p.Steps[i]
		s.Text += " changed"
		p.Steps[i] = &s
		if guardOPCCorpusNoopCase(opcCorpusNoopCaseID, &p, 8) == nil {
			t.Fatalf("accepted step %d drift", i)
		}
	}
	if guardOPCCorpusNoopCase(opcCorpusNoopCaseID, picked, 9) == nil {
		t.Fatal("accepted line drift")
	}
}

type corpusFile struct {
	ID      string `json:"id"`
	Path    string `json:"path"`
	Role    string `json:"role"`
	Origins []struct {
		Repository string `json:"repository"`
		Kind       string `json:"kind"`
	} `json:"origins"`
	Aliases []string `json:"aliases"`
}

type corpusEntry struct{ id, path string }

// Enumerate provenance in the pinned manifest, not a consumer-maintained list or native label map.
func corpusEntries(files []corpusFile) (goFiles, pythonFiles []corpusEntry, err error) {
	goIDs, pyIDs := map[string]bool{}, map[string]bool{}
	paths := map[string]bool{}
	for _, f := range files {
		if f.Role != "fixture" || !strings.HasPrefix(f.Path, "fixtures/") || !(strings.HasSuffix(f.Path, ".docx") || strings.HasSuffix(f.Path, ".xlsx") || strings.HasSuffix(f.Path, ".pptx")) {
			continue
		}
		directGo, directPython, derived := false, false, false
		for _, o := range f.Origins {
			if o.Repository == goOrigin && o.Kind == "" {
				directGo = true
			}
			if o.Repository == pythonOrigin {
				if o.Kind == "derived-from" {
					derived = true
				} else if o.Kind == "" {
					directPython = true
				}
			}
		}
		pythonTestdata := false
		for _, a := range f.Aliases {
			if strings.HasPrefix(a, pythonTestdataAlias) {
				pythonTestdata = true
			}
		}
		if directGo && (directPython || derived) {
			return nil, nil, fmt.Errorf("mixed provenance: %s", f.ID)
		}
		if directGo {
			if goIDs[f.ID] {
				return nil, nil, fmt.Errorf("duplicate Go ID %s", f.ID)
			}
			if paths[f.Path] {
				return nil, nil, fmt.Errorf("duplicate corpus path %s", f.Path)
			}
			goIDs[f.ID], paths[f.Path] = true, true
			goFiles = append(goFiles, corpusEntry{f.ID, f.Path})
		}
		if directPython && pythonTestdata && !derived {
			if pyIDs[f.ID] || goIDs[f.ID] {
				return nil, nil, fmt.Errorf("duplicate Python or shared ID %s", f.ID)
			}
			if paths[f.Path] {
				return nil, nil, fmt.Errorf("duplicate corpus path %s", f.Path)
			}
			pyIDs[f.ID], paths[f.Path] = true, true
			pythonFiles = append(pythonFiles, corpusEntry{f.ID, f.Path})
		}
	}
	if len(goFiles) != 36 || len(pythonFiles) != 35 {
		return nil, nil, fmt.Errorf("corpus drift: Go %d/36 Python %d/35", len(goFiles), len(pythonFiles))
	}
	for id := range goIDs {
		if pyIDs[id] {
			return nil, nil, fmt.Errorf("overlapping corpus ID %s", id)
		}
	}
	if len(goIDs)+len(pyIDs) != 71 {
		return nil, nil, fmt.Errorf("total corpus IDs %d/71", len(goIDs)+len(pyIDs))
	}
	return goFiles, pythonFiles, nil
}

func readCorpusEntries() ([]corpusEntry, []corpusEntry, error) {
	b, err := os.ReadFile(testutil.ReferencePath("manifest.json"))
	if err != nil {
		return nil, nil, err
	}
	var m struct {
		Schema int          `json:"schemaVersion"`
		Files  []corpusFile `json:"files"`
	}
	if err = json.Unmarshal(b, &m); err != nil {
		return nil, nil, err
	}
	if m.Schema != 2 {
		return nil, nil, fmt.Errorf("fixture manifest schema %d", m.Schema)
	}
	return corpusEntries(m.Files)
}

var corpusLimits = packaging.Limits{MaxSourceBytes: 64 << 20, MaxEntries: 4096, MaxPartBytes: 32 << 20, MaxTotalBytes: 128 << 20}

// The public operation is injected only so the acceptance test can prove that it
// detects a dropped write or corrupted emitted bytes; normal execution uses the real API.
func corpusRoundTrip(source []byte, open func([]byte) (*packaging.Preserved, error), emit func(*packaging.Preserved) ([]byte, error)) error {
	first, err := open(source)
	if err != nil || first == nil {
		return fmt.Errorf("first intake: %v (nil=%t)", err, first == nil)
	}
	output, err := emit(first)
	if err != nil {
		return err
	}
	if !bytes.Equal(output, source) {
		return fmt.Errorf("first no-op archive differs")
	}
	second, err := open(output)
	if err != nil || second == nil {
		return fmt.Errorf("reopen: %v (nil=%t)", err, second == nil)
	}
	reopened, err := emit(second)
	if err != nil {
		return err
	}
	if !bytes.Equal(reopened, source) {
		return fmt.Errorf("reopened no-op archive differs")
	}
	return nil
}

func realCorpusOpen(b []byte) (*packaging.Preserved, error) {
	return packaging.OpenPreserved(b, corpusLimits)
}
func realCorpusEmit(p *packaging.Preserved) ([]byte, error) {
	var b bytes.Buffer
	if err := p.WriteTo(&b); err != nil {
		return nil, err
	}
	return bytes.Clone(b.Bytes()), nil
}

// Count only successful public open/write/reopen/write cycles, never attempted rows.
func visitCorpus(entries []corpusEntry, load func(corpusEntry) ([]byte, error), roundTrip func([]byte) error) (int, error) {
	successes := 0
	for _, e := range entries {
		source, err := load(e)
		if err != nil {
			return successes, fmt.Errorf("%s: %w", e.path, err)
		}
		if err := roundTrip(source); err != nil {
			return successes, fmt.Errorf("%s: %w", e.path, err)
		}
		successes++
	}
	return successes, nil
}

func loadCorpusFixture(e corpusEntry) ([]byte, error) {
	path, err := testutil.LookupFixture(e.id)
	if err != nil {
		return nil, err
	}
	// Independently check the manifest path matches the selected ID's actual location.
	if filepath.Clean(path) != filepath.Clean(testutil.ReferencePath(filepath.FromSlash(e.path))) {
		return nil, fmt.Errorf("manifest ID/path mismatch: %s", e.id)
	}
	return os.ReadFile(path)
}

func TestOPCCorpusNoopNegativeControls(t *testing.T) {
	goFiles, pyFiles, err := readCorpusEntries()
	if err != nil {
		t.Fatal(err)
	}
	b, err := os.ReadFile(testutil.ReferencePath("manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var m struct {
		Files []corpusFile `json:"files"`
	}
	if err := json.Unmarshal(b, &m); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name  string
		files []corpusFile
	}{
		{"missing", func() []corpusFile {
			var out []corpusFile
			for _, f := range m.Files {
				if f.ID != goFiles[0].id {
					out = append(out, f)
				}
			}
			return out
		}()},
		{"duplicate", func() []corpusFile {
			var out []corpusFile
			for _, f := range m.Files {
				out = append(out, f)
				if f.ID == goFiles[0].id {
					out = append(out, f)
				}
			}
			return out
		}()},
	} {
		t.Run(tc.name, func(t *testing.T) {
			// A missing/duplicate selected fixture must never preserve a passing 71-row corpus.
			if _, _, err := corpusEntries(tc.files); err == nil {
				t.Fatal("accepted corpus selection drift")
			}
		})
	}
	if goFiles[0].id == pyFiles[0].id {
		t.Fatal("corpus IDs overlap")
	}
	if n, err := visitCorpus(goFiles, func(e corpusEntry) ([]byte, error) {
		if e.id == goFiles[0].id {
			return nil, fmt.Errorf("missing fixture")
		}
		return nil, nil
	}, func([]byte) error { return nil }); err == nil || n != 0 {
		t.Fatal("omitted fixture was counted")
	}
	n, err := visitCorpus(goFiles[:len(goFiles)-1], loadCorpusFixture, func(b []byte) error { return corpusRoundTrip(b, realCorpusOpen, realCorpusEmit) })
	if err != nil || n != 35 {
		t.Fatalf("short traversal: %d, %v", n, err)
	}
	if err := requireCorpusCounts(n, 35); err == nil {
		t.Fatal("accepted one dropped Go iteration")
	}
	if err := corpusRoundTrip([]byte("not a ZIP"), realCorpusOpen, realCorpusEmit); err == nil {
		t.Fatal("accepted invalid archive")
	}
	source, err := loadCorpusFixture(pyFiles[0])
	if err != nil {
		t.Fatal(err)
	}
	for _, stage := range []int{1, 2} {
		calls := 0
		emit := func(p *packaging.Preserved) ([]byte, error) {
			b, err := realCorpusEmit(p)
			calls++
			if err == nil && calls == stage {
				b = append(bytes.Clone(b), 0)
			}
			return b, err
		}
		if err := corpusRoundTrip(source, realCorpusOpen, emit); err == nil {
			t.Fatalf("accepted stage %d output corruption", stage)
		}
	}
}

func requireCorpusCounts(goCount, pyCount int) error {
	if goCount != 36 || pyCount != 35 || goCount+pyCount != 71 {
		return fmt.Errorf("successful corpus counts Go %d Python %d total %d", goCount, pyCount, goCount+pyCount)
	}
	return nil
}

func opcCorpusNoopSteps(sc *godog.ScenarioContext) {
	var goFiles, pyFiles []corpusEntry
	var goCount, pyCount int
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		goFiles, pyFiles = nil, nil
		goCount, pyCount = 0, 0
		return ctx, nil
	})
	sc.Step(`^the go-ooxml and python-office-mcp-server fixture corpora are enumerated$`, func() error {
		var err error
		goFiles, pyFiles, err = readCorpusEntries()
		return err
	})
	sc.Step(`^each OOXML fixture package is opened and serialized without edits through the OPC layer$`, func() error {
		if len(goFiles) != 36 || len(pyFiles) != 35 {
			return fmt.Errorf("corpus not enumerated")
		}
		var err error
		goCount, err = visitCorpus(goFiles, loadCorpusFixture, func(b []byte) error { return corpusRoundTrip(b, realCorpusOpen, realCorpusEmit) })
		if err != nil {
			return err
		}
		pyCount, err = visitCorpus(pyFiles, loadCorpusFixture, func(b []byte) error { return corpusRoundTrip(b, realCorpusOpen, realCorpusEmit) })
		return err
	})
	sc.Step(`^every reopened package matches its original whole-archive bytes$`, func() error {
		if goCount != len(goFiles) || pyCount != len(pyFiles) {
			return fmt.Errorf("successful reopen counts %d/%d, want %d/%d", goCount, pyCount, len(goFiles), len(pyFiles))
		}
		return nil
	})
	sc.Step(`^both fixture corpora contribute their exact known nonzero fixture counts$`, func() error {
		return requireCorpusCounts(goCount, pyCount)
	})
}
