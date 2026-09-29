package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const xmlEntityValuesCaseID = "@id-xml-entity-values"
const xmlEntitySourceJSON = `"<r a=\"&quot;&apos;\">&#x41;&#65;&amp;&lt;&gt;</r>"`
const xmlEntityAttributeJSON = `"\"'"`
const xmlEntityTextJSON = `"AA&<>"`

func guardXMLEntityValuesCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"XML values input encoded as JSON " + xmlEntitySourceJSON,
		"the XML values input is parsed",
		"the root attribute a equals JSON " + xmlEntityAttributeJSON,
		"the root text equals JSON " + xmlEntityTextJSON,
	}
	if id != xmlEntityValuesCaseID || line != 45 || p.Name != "Decode predefined entities and decimal and hexadecimal references" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected XML entity-values case %s %q at %d", id, p.Name, line)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("XML entity-values step %d drift", i+1)
		}
	}
	return nil
}

// The canonical scenario lives under this Rule, and has no scenario profile.
func guardXMLEntityValuesRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "XML parsing and value inspection" {
		return fmt.Errorf("XML entity-values feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil {
			continue
		}
		if child.Rule.Name != "XML values, namespace lookup and safe escaping" {
			continue
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("XML entity-values rule gained background")
			}
			if member.Scenario == nil {
				continue
			}
			for _, tag := range member.Scenario.Tags {
				if tag.Name == xmlEntityValuesCaseID {
					if len(member.Scenario.Tags) != 1 || int(tag.Location.Line) != 44 || len(member.Scenario.Examples) != 0 || len(member.Scenario.Steps) != 4 || int(member.Scenario.Location.Line) != 45 {
						return fmt.Errorf("XML entity-values canonical scenario structure drift")
					}
					found++
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("XML entity-values rule owns %d canonical scenarios", found)
	}
	return nil
}

func TestXMLEntityValuesGuardRejectsDrift(t *testing.T) {
	path := xmlParsingFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil || guardXMLEntityValuesRule(doc) != nil {
		t.Fatalf("canonical XML entity-values Rule drift: %v", err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == xmlEntityValuesCaseID {
				if selected != nil {
					t.Fatal("duplicate XML entity-values pickle")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardXMLEntityValuesCase(xmlEntityValuesCaseID, selected, 45) != nil {
		t.Fatal("canonical XML entity-values case guard failed")
	}
	for _, tc := range []struct {
		name   string
		mutate func(*messages.Pickle)
	}{
		{"source", func(p *messages.Pickle) { p.Steps[0].Text += " changed" }},
		{"attribute", func(p *messages.Pickle) { p.Steps[2].Text += " changed" }},
		{"text", func(p *messages.Pickle) { p.Steps[3].Text += " changed" }},
		{"scenario name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "unexpected") }},
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
			tc.mutate(&clone)
			if guardXMLEntityValuesCase(xmlEntityValuesCaseID, &clone, 45) == nil {
				t.Fatal("guard accepted XML entity-values drift")
			}
		})
	}
	if guardXMLEntityValuesCase("@id-xml-stylesheet-processing-instruction", selected, 45) == nil || guardXMLEntityValuesCase(xmlEntityValuesCaseID, selected, 46) == nil {
		t.Fatal("guard accepted XML entity-values ID or line drift")
	}
}

func xmlEntityValuesSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc = nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^XML values input encoded as JSON (".*")$`, func(raw string) error {
		if raw != xmlEntitySourceJSON && raw != xmlStylesheetPISourceJSON && raw != xmlImplicitPrefixSourceJSON && raw != xmlExpandedAttributeSourceJSON {
			return fmt.Errorf("XML values source drift")
		}
		var source string
		if err := json.Unmarshal([]byte(raw), &source); err != nil {
			return err
		}
		caller = []byte(source)
		original = bytes.Clone(caller)
		if !bytes.Equal(caller, []byte(`<r a="&quot;&apos;">&#x41;&#65;&amp;&lt;&gt;</r>`)) && !bytes.Equal(caller, []byte(`<?xml-stylesheet href="style.xsl"?><r/>`)) && !bytes.Equal(caller, []byte(xmlImplicitPrefixInput)) && !bytes.Equal(caller, []byte(xmlExpandedAttributeInput)) {
			return fmt.Errorf("XML values decoded source drift")
		}
		return nil
	})
	sc.Step(`^the XML values input is parsed$`, func() error {
		if caller == nil || !bytes.Equal(caller, original) {
			return fmt.Errorf("missing caller-owned XML values source")
		}
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil || !bytes.Equal(caller, original) {
			return fmt.Errorf("XML entity-values parse/caller drift: %v", err)
		}
		return nil
	})
	sc.Step(`^the root attribute a equals JSON (".*")$`, func(raw string) error {
		var expected string
		if err := json.Unmarshal([]byte(raw), &expected); err != nil {
			return err
		}
		if raw != xmlEntityAttributeJSON || expected != "\"'" || doc == nil {
			return fmt.Errorf("XML entity-values attribute expectation drift")
		}
		es := doc.Elements()
		if len(es) != 1 || es[0].Name() != (xml.Name{Local: "r"}) {
			return fmt.Errorf("XML entity-values root name/count drift")
		}
		attrs := es[0].Attributes()
		if len(attrs) != 1 || attrs[0].Name != (xml.Name{Local: "a"}) || attrs[0].Value != expected || !bytes.Equal(caller, original) {
			return fmt.Errorf("XML entity-values decoded attribute or caller drift")
		}
		return nil
	})
	sc.Step(`^the root text equals JSON (".*")$`, func(raw string) error {
		var expected string
		if err := json.Unmarshal([]byte(raw), &expected); err != nil {
			return err
		}
		if raw != xmlEntityTextJSON || expected != "AA&<>" || doc == nil {
			return fmt.Errorf("XML entity-values text expectation drift")
		}
		es := doc.Elements()
		if len(es) != 1 || es[0].Name() != (xml.Name{Local: "r"}) {
			return fmt.Errorf("XML entity-values root changed")
		}
		text, leaf := es[0].Text()
		if !leaf || text != expected || !bytes.Equal(caller, original) {
			return fmt.Errorf("XML entity-values decoded leaf text or caller drift: %q", text)
		}
		var err error
		output, err = doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(output, original) || !bytes.Equal(caller, original) {
			return fmt.Errorf("XML entity-values empty edit changed source: %v", err)
		}
		output[0] = '!'
		again, err := doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(again, original) || !bytes.Equal(caller, original) {
			return fmt.Errorf("XML entity-values no-op return aliases source or caller: %v", err)
		}
		bad := []byte(`<r>&bogus;</r>`)
		badBefore := bytes.Clone(bad)
		invalid, err := losslessxml.Parse(bad)
		if err == nil || invalid != nil || !bytes.Equal(bad, badBefore) {
			return fmt.Errorf("XML entity-values unknown-entity refusal or caller custody drift: %v", err)
		}
		valid, err := losslessxml.Parse(caller)
		if err != nil || len(valid.Elements()) != 1 {
			return fmt.Errorf("XML entity-values valid retry failed: %v", err)
		}
		got, ok := valid.Elements()[0].Text()
		if !ok || got != expected || !bytes.Equal(caller, original) {
			return fmt.Errorf("XML entity-values retry/caller drift")
		}
		return nil
	})
	sc.Step(`^the root qualified name equals r$`, func() error {
		return assertStylesheetPI(caller, original, doc)
	})
	registerXMLImplicitPrefixSteps(sc, &caller, &original, &doc)
	registerXMLExpandedAttributeSteps(sc, &caller, &original, &doc)
}
