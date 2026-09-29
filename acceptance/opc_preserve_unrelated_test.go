package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const opcPreserveUnrelatedCaseID = "@id-opc-package-preserve-unrelated"
const opcMainPart = "word/document.xml"
const opcOpaquePart = "custom/opaque.bin"
const opcMainAlpha = `<document>Alpha</document>`
const opcMainBeta = `<document>Beta</document>`

var opcOpaqueBytes = []byte{0, 42, 128, 255, 10, 0}

var opcPreserveSteps = []string{
	"a valid OPC package with a main XML part and an unrelated binary payload",
	"the main XML part text is changed and the package is reopened",
	"the edited part contains the new text after reopen",
	"the unrelated payload bytes remain unchanged",
}

func guardOPCPreserveUnrelatedCase(id string, p *messages.Pickle, line int) error {
	if id != opcPreserveUnrelatedCaseID || line != 21 || p.Name != "Changing one part preserves unrelated payload after reopen" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(opcPreserveSteps) {
		return fmt.Errorf("OPC preserve-unrelated case drift: %s %q at %d", id, p.Name, line)
	}
	for i, want := range opcPreserveSteps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("OPC preserve-unrelated step %d drift", i+1)
		}
	}
	return nil
}

func guardOPCPreserveUnrelatedRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "OPC package custody, transactions and save destinations" || len(doc.Feature.Tags) != 1 || doc.Feature.Tags[0].Name != "@planned" {
		return fmt.Errorf("OPC preservation feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "OPC package custody and transactional part edits" {
			continue
		}
		if len(child.Rule.Tags) != 0 {
			return fmt.Errorf("OPC preservation rule gained tags")
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("OPC preservation rule gained background")
			}
			if member.Scenario == nil {
				continue
			}
			for _, tag := range member.Scenario.Tags {
				if tag.Name == opcPreserveUnrelatedCaseID {
					if len(member.Scenario.Tags) != 1 || int(tag.Location.Line) != 20 || int(member.Scenario.Location.Line) != 21 || len(member.Scenario.Examples) != 0 || len(member.Scenario.Steps) != 4 {
						return fmt.Errorf("OPC preserve-unrelated structure drift")
					}
					found++
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("OPC preservation rule owns %d selected scenarios", found)
	}
	return nil
}

func TestOPCPreserveUnrelatedGuardRejectsDrift(t *testing.T) {
	path := packagePreservationFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil || guardOPCPreserveUnrelatedRule(doc) != nil {
		t.Fatalf("OPC preservation Rule drift: %v", err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == opcPreserveUnrelatedCaseID {
				if selected != nil {
					t.Fatal("duplicate OPC preserve-unrelated pickle")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardOPCPreserveUnrelatedCase(opcPreserveUnrelatedCaseID, selected, 21) != nil {
		t.Fatal("OPC preserve-unrelated case guard failed")
	}
	for _, tc := range []struct {
		name   string
		change func(*messages.Pickle)
	}{
		{"given", func(p *messages.Pickle) { p.Steps[0].Text += " changed" }},
		{"when", func(p *messages.Pickle) { p.Steps[1].Text += " changed" }},
		{"then", func(p *messages.Pickle) { p.Steps[2].Text += " changed" }},
		{"and", func(p *messages.Pickle) { p.Steps[3].Text += " changed" }},
		{"name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "extra") }},
		{"argument", func(p *messages.Pickle) { p.Steps[1].Argument = &messages.PickleStepArgument{} }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			clone := *selected
			clone.AstNodeIds = append([]string(nil), selected.AstNodeIds...)
			clone.Steps = append([]*messages.PickleStep(nil), selected.Steps...)
			for i, step := range clone.Steps {
				copyStep := *step
				clone.Steps[i] = &copyStep
			}
			tc.change(&clone)
			if guardOPCPreserveUnrelatedCase(opcPreserveUnrelatedCaseID, &clone, 21) == nil {
				t.Fatal("guard accepted OPC preservation drift")
			}
		})
	}
	if guardOPCPreserveUnrelatedCase("@id-opc-package-corpus-noop", selected, 21) == nil || guardOPCPreserveUnrelatedCase(opcPreserveUnrelatedCaseID, selected, 22) == nil {
		t.Fatal("guard accepted ID or line drift")
	}
}

// The source fixture is built by archive/zip, independently of the preserved
// package writer and of the bytes inspected after its edit.
func opcPreserveSource() ([]byte, error) {
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	for _, entry := range []struct {
		name string
		data []byte
	}{
		{"[Content_Types].xml", []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)},
		{"_rels/.rels", []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)},
		{opcMainPart, []byte(opcMainAlpha)},
		{opcOpaquePart, bytes.Clone(opcOpaqueBytes)},
	} {
		w, err := z.Create(entry.name)
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(entry.data); err != nil {
			return nil, err
		}
	}
	if err := z.Close(); err != nil {
		return nil, err
	}
	return bytes.Clone(b.Bytes()), nil
}

func opcZipMembers(archive []byte) (map[string][]byte, error) {
	z, err := zip.NewReader(bytes.NewReader(archive), int64(len(archive)))
	if err != nil {
		return nil, err
	}
	out := map[string][]byte{}
	for _, f := range z.File {
		if _, ok := out[f.Name]; ok {
			return nil, fmt.Errorf("duplicate ZIP member %s", f.Name)
		}
		r, err := f.Open()
		if err != nil {
			return nil, err
		}
		data, readErr := io.ReadAll(r)
		closeErr := r.Close()
		if readErr != nil {
			return nil, readErr
		}
		if closeErr != nil {
			return nil, closeErr
		}
		out[f.Name] = data
	}
	return out, nil
}

func opcText(data []byte) (string, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	depth, roots := 0, 0
	var text bytes.Buffer
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			if depth != 0 || roots != 0 || v.Name != (xml.Name{Local: "document"}) || len(v.Attr) != 0 {
				return "", fmt.Errorf("unexpected main XML element")
			}
			depth, roots = 1, 1
		case xml.CharData:
			if depth != 1 {
				return "", fmt.Errorf("unexpected main XML text")
			}
			text.Write(v)
		case xml.EndElement:
			if depth != 1 || v.Name != (xml.Name{Local: "document"}) {
				return "", fmt.Errorf("unexpected main XML end")
			}
			depth = 0
		}
	}
	if depth != 0 || roots != 1 || text.Len() == 0 {
		return "", fmt.Errorf("incomplete main XML")
	}
	return text.String(), nil
}

func opcPreserveUnrelatedSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var sourceMembers, savedMembers map[string][]byte
	var reopened *packaging.Preserved
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output = nil, nil, nil
		sourceMembers, savedMembers, reopened = nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^a valid OPC package with a main XML part and an unrelated binary payload$`, func() error {
		var err error
		caller, err = opcPreserveSource()
		if err != nil {
			return err
		}
		original = bytes.Clone(caller)
		sourceMembers, err = opcZipMembers(caller)
		if err != nil || len(sourceMembers) != 4 || !bytes.Equal(sourceMembers[opcOpaquePart], opcOpaqueBytes) {
			return fmt.Errorf("independent OPC fixture mismatch: %v", err)
		}
		text, err := opcText(sourceMembers[opcMainPart])
		if err != nil || text != "Alpha" {
			return fmt.Errorf("initial main text mismatch: %v", err)
		}
		return nil
	})
	sc.Step(`^the main XML part text is changed and the package is reopened$`, func() error {
		if !bytes.Equal(caller, original) || sourceMembers == nil {
			return fmt.Errorf("caller/source missing or changed")
		}
		p, err := packaging.OpenPreserved(caller, packaging.Limits{})
		if err != nil || p == nil {
			return fmt.Errorf("source intake: %v", err)
		}
		main, fingerprint, err := p.Part(opcMainPart)
		if err != nil || !bytes.Equal(main, sourceMembers[opcMainPart]) {
			return fmt.Errorf("main part lookup: %v", err)
		}
		if err := p.Replace([]packaging.Replacement{{Part: opcMainPart, ExpectedSHA256: fingerprint, Data: []byte(opcMainBeta)}}); err != nil {
			return err
		}
		var b bytes.Buffer
		if err := p.WriteTo(&b); err != nil {
			return err
		}
		output = bytes.Clone(b.Bytes())
		reopened, err = packaging.OpenPreserved(output, packaging.Limits{})
		if err != nil || reopened == nil || !bytes.Equal(caller, original) {
			return fmt.Errorf("edited archive reopen/caller custody: %v", err)
		}
		return nil
	})
	sc.Step(`^the edited part contains the new text after reopen$`, func() error {
		if reopened == nil {
			return fmt.Errorf("archive not reopened")
		}
		var err error
		savedMembers, err = opcZipMembers(output)
		if err != nil || len(savedMembers) != len(sourceMembers) {
			return fmt.Errorf("independent output ZIP read/inventory: %v", err)
		}
		part, _, err := reopened.Part(opcMainPart)
		if err != nil || !bytes.Equal(part, savedMembers[opcMainPart]) {
			return fmt.Errorf("reopened main member differs from ZIP reader: %v", err)
		}
		text, err := opcText(part)
		if err != nil || text != "Beta" || !bytes.Equal(caller, original) {
			return fmt.Errorf("edited main text/caller custody: %q %v", text, err)
		}
		return nil
	})
	sc.Step(`^the unrelated payload bytes remain unchanged$`, func() error {
		if savedMembers == nil || reopened == nil || len(savedMembers) != len(sourceMembers) {
			return fmt.Errorf("output not checked")
		}
		part, _, err := reopened.Part(opcOpaquePart)
		if err != nil || !bytes.Equal(part, savedMembers[opcOpaquePart]) || !bytes.Equal(savedMembers[opcOpaquePart], sourceMembers[opcOpaquePart]) || !bytes.Equal(caller, original) {
			return fmt.Errorf("unrelated binary or caller bytes changed: %v", err)
		}
		for _, name := range []string{"[Content_Types].xml", "_rels/.rels"} {
			if !bytes.Equal(savedMembers[name], sourceMembers[name]) {
				return fmt.Errorf("unrelated OPC registry member changed: %s", name)
			}
		}
		return nil
	})
}

func TestOPCPreserveUnrelatedStaleFingerprintAndRetry(t *testing.T) {
	caller, err := opcPreserveSource()
	if err != nil {
		t.Fatal(err)
	}
	original := bytes.Clone(caller)
	p, err := packaging.OpenPreserved(caller, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	if err := p.Replace([]packaging.Replacement{{Part: opcMainPart, ExpectedSHA256: "stale", Data: []byte(opcMainBeta)}}); err == nil {
		t.Fatal("stale fingerprint accepted")
	} else {
		var refusal *packaging.Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "stale_target" {
			t.Fatalf("unexpected refusal: %v", err)
		}
	}
	var noOp bytes.Buffer
	if err := p.WriteTo(&noOp); err != nil || !bytes.Equal(noOp.Bytes(), original) || !bytes.Equal(caller, original) {
		t.Fatalf("failed replacement changed session/source: %v", err)
	}
	_, fp, err := p.Part(opcMainPart)
	if err != nil {
		t.Fatal(err)
	}
	if err := p.Replace([]packaging.Replacement{{Part: opcMainPart, ExpectedSHA256: fp, Data: []byte(opcMainBeta)}}); err != nil {
		t.Fatal(err)
	}
	var b bytes.Buffer
	if err := p.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	q, err := packaging.OpenPreserved(b.Bytes(), packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	part, _, err := q.Part(opcMainPart)
	if err != nil {
		t.Fatal(err)
	}
	text, err := opcText(part)
	if err != nil || text != "Beta" || !bytes.Equal(caller, original) {
		t.Fatalf("valid retry failed: %q %v", text, err)
	}
}
