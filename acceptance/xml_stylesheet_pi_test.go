package acceptance

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const xmlStylesheetPICaseID = "@id-xml-stylesheet-processing-instruction"
const xmlStylesheetPISourceJSON = `"<?xml-stylesheet href=\"style.xsl\"?><r/>"`
const xmlStylesheetPIInput = `<?xml-stylesheet href="style.xsl"?><r/>`
const xmlStylesheetPIRoot = `<r/>`

func guardXMLStylesheetPICase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"XML values input encoded as JSON " + xmlStylesheetPISourceJSON,
		"the XML values input is parsed",
		"the root qualified name equals r",
	}
	if id != xmlStylesheetPICaseID || line != 52 || p.Name != "Accept a stylesheet processing instruction before the root" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected XML stylesheet PI case %s %q at %d", id, p.Name, line)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("XML stylesheet PI step %d drift", i+1)
		}
	}
	return nil
}

func guardXMLStylesheetPIRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "XML parsing and value inspection" {
		return fmt.Errorf("XML stylesheet PI feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "XML values, namespace lookup and safe escaping" {
			continue
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("XML stylesheet PI rule gained background")
			}
			if member.Scenario == nil {
				continue
			}
			for _, tag := range member.Scenario.Tags {
				if tag.Name == xmlStylesheetPICaseID {
					if len(member.Scenario.Tags) != 1 || int(tag.Location.Line) != 51 || len(member.Scenario.Examples) != 0 || len(member.Scenario.Steps) != 3 || int(member.Scenario.Location.Line) != 52 {
						return fmt.Errorf("XML stylesheet PI canonical structure drift")
					}
					found++
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("XML stylesheet PI rule owns %d canonical scenarios", found)
	}
	return nil
}

func TestXMLStylesheetPIGuardRejectsDrift(t *testing.T) {
	path := xmlParsingFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil || guardXMLStylesheetPIRule(doc) != nil {
		t.Fatalf("canonical XML stylesheet PI Rule drift: %v", err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == xmlStylesheetPICaseID {
				if selected != nil {
					t.Fatal("duplicate XML stylesheet PI pickle")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardXMLStylesheetPICase(xmlStylesheetPICaseID, selected, 52) != nil {
		t.Fatal("canonical XML stylesheet PI case guard failed")
	}
	for _, tc := range []struct {
		name   string
		mutate func(*messages.Pickle)
	}{
		{"source", func(p *messages.Pickle) { p.Steps[0].Text += " changed" }},
		{"parse", func(p *messages.Pickle) { p.Steps[1].Text += " changed" }},
		{"root", func(p *messages.Pickle) { p.Steps[2].Text += " changed" }},
		{"scenario name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "unexpected") }},
		{"argument", func(p *messages.Pickle) { p.Steps[0].Argument = &messages.PickleStepArgument{} }},
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
			if guardXMLStylesheetPICase(xmlStylesheetPICaseID, &clone, 52) == nil {
				t.Fatal("guard accepted XML stylesheet PI drift")
			}
		})
	}
	if guardXMLStylesheetPICase(xmlEntityValuesCaseID, selected, 52) == nil || guardXMLStylesheetPICase(xmlStylesheetPICaseID, selected, 51) == nil || guardXMLStylesheetPICase(xmlStylesheetPICaseID, selected, 45) == nil {
		t.Fatal("guard accepted XML stylesheet PI ID or line drift")
	}
}

func assertStylesheetPIRoot(doc *losslessxml.Document) error {
	if doc == nil {
		return fmt.Errorf("XML stylesheet PI missing parsed document")
	}
	elements := doc.Elements()
	if len(elements) != 1 || elements[0].Name() != (xml.Name{Local: "r"}) {
		return fmt.Errorf("XML stylesheet PI root count or qualified name drift")
	}
	if _, ok := elements[0].Parent(); ok || !bytes.Equal(elements[0].Raw(), []byte(xmlStylesheetPIRoot)) {
		return fmt.Errorf("XML stylesheet PI root ownership or raw markup drift")
	}
	return nil
}

func assertStylesheetPI(caller, original []byte, doc *losslessxml.Document) error {
	if !bytes.Equal(original, []byte(xmlStylesheetPIInput)) || !bytes.Equal(caller, original) {
		return fmt.Errorf("XML stylesheet PI source or caller drift")
	}
	if err := assertStylesheetPIRoot(doc); err != nil {
		return err
	}
	start, end := doc.Elements()[0].SourceRange()
	prefix := []byte(`<?xml-stylesheet href="style.xsl"?>`)
	if start != len(prefix) || end != len(original) || !bytes.Equal(original[:start], prefix) || !bytes.Equal(original[start:end], []byte(xmlStylesheetPIRoot)) {
		return fmt.Errorf("XML stylesheet PI source prefix or root range drift")
	}
	output, err := doc.Edit(nil, nil)
	if err != nil || !bytes.Equal(output, original) || !bytes.Equal(caller, original) {
		return fmt.Errorf("XML stylesheet PI empty edit/caller drift: %v", err)
	}
	output[0] = '!'
	again, err := doc.Edit(nil, nil)
	if err != nil || !bytes.Equal(again, original) || !bytes.Equal(caller, original) {
		return fmt.Errorf("XML stylesheet PI output aliases source or caller: %v", err)
	}
	caller[0] = '!'
	third, err := doc.Edit(nil, nil)
	if err != nil || !bytes.Equal(third, original) {
		return fmt.Errorf("XML stylesheet PI snapshot aliases caller: %v", err)
	}
	return nil
}

func TestXMLStylesheetPINegativeControls(t *testing.T) {
	for _, source := range [][]byte{
		[]byte(`<?xml-stylesheet href="style.xsl"<r/>`),
		[]byte(`<?xml-stylesheet href="style.xsl"?><r/><?broken`),
	} {
		caller := bytes.Clone(source)
		result, err := losslessxml.Parse(caller)
		if err == nil || result != nil || !bytes.Equal(caller, source) {
			t.Fatalf("broken PI accepted or changed caller: %q: %v", source, err)
		}
	}
	wrong := []byte(`<?xml-stylesheet href="style.xsl"?><s/>`)
	before := bytes.Clone(wrong)
	doc, err := losslessxml.Parse(wrong)
	if err != nil || !bytes.Equal(wrong, before) {
		t.Fatalf("wrong-root control failed to parse or changed caller: %v", err)
	}
	if assertStylesheetPIRoot(doc) == nil {
		t.Fatal("wrong-root control accepted s as r")
	}
}
