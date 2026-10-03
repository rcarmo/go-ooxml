package acceptance

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const xmlImplicitPrefixCaseID = "@id-xml-implicit-xml-prefix"
const xmlImplicitPrefixSourceJSON = `"<r xml:lang=\"en\"/>"`
const xmlImplicitPrefixInput = `<r xml:lang="en"/>`
const xmlImplicitPrefixURI = "http://www.w3.org/XML/1998/namespace"

func xmlImplicitPrefixLine() int {
	if xmlLexicalCandidate() || historicalAPI18Shape() {
		return 99
	}
	return 72
}

func guardXMLImplicitPrefixCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"XML values input encoded as JSON " + xmlImplicitPrefixSourceJSON,
		"the XML values input is parsed",
		`the root namespace URI equals JSON ""`,
		`the root attribute xml:lang equals JSON "en"`,
		"the implicit xml namespace URI is " + xmlImplicitPrefixURI,
	}
	if id != xmlImplicitPrefixCaseID || line != xmlImplicitPrefixLine() || p.Name != "The xml prefix is bound without a namespace declaration" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected implicit xml prefix case %s %q at %d", id, p.Name, line)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("implicit xml prefix step %d drift", i+1)
		}
	}
	return nil
}

func guardXMLImplicitPrefixRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "XML parsing and value inspection" {
		return fmt.Errorf("implicit xml prefix feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "XML values, namespace lookup and safe escaping" {
			continue
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("implicit xml prefix rule gained background")
			}
			if member.Scenario == nil {
				continue
			}
			for _, tag := range member.Scenario.Tags {
				if tag.Name == xmlImplicitPrefixCaseID {
					if len(member.Scenario.Tags) != 1 || int(tag.Location.Line) != xmlImplicitPrefixLine()-1 || len(member.Scenario.Examples) != 0 || len(member.Scenario.Steps) != 5 || int(member.Scenario.Location.Line) != xmlImplicitPrefixLine() {
						return fmt.Errorf("implicit xml prefix canonical structure drift")
					}
					found++
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("implicit xml prefix rule owns %d canonical scenarios", found)
	}
	return nil
}

func TestXMLImplicitPrefixGuardRejectsDrift(t *testing.T) {
	path := xmlParsingFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil || guardXMLImplicitPrefixRule(doc) != nil {
		t.Fatalf("canonical implicit xml prefix Rule drift: %v", err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == xmlImplicitPrefixCaseID {
				if selected != nil {
					t.Fatal("duplicate implicit xml prefix pickle")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardXMLImplicitPrefixCase(xmlImplicitPrefixCaseID, selected, xmlImplicitPrefixLine()) != nil {
		t.Fatal("canonical implicit xml prefix case guard failed")
	}
	for _, tc := range []struct {
		name   string
		mutate func(*messages.Pickle)
	}{
		{"source", func(p *messages.Pickle) { p.Steps[0].Text += " changed" }},
		{"parse", func(p *messages.Pickle) { p.Steps[1].Text += " changed" }},
		{"root", func(p *messages.Pickle) { p.Steps[2].Text += " changed" }},
		{"attribute", func(p *messages.Pickle) { p.Steps[3].Text += " changed" }},
		{"binding", func(p *messages.Pickle) { p.Steps[4].Text += " changed" }},
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
			if guardXMLImplicitPrefixCase(xmlImplicitPrefixCaseID, &clone, xmlImplicitPrefixLine()) == nil {
				t.Fatal("guard accepted implicit xml prefix drift")
			}
		})
	}
	if guardXMLImplicitPrefixCase(xmlEntityValuesCaseID, selected, xmlImplicitPrefixLine()) == nil || guardXMLImplicitPrefixCase(xmlImplicitPrefixCaseID, selected, xmlImplicitPrefixLine()-1) == nil || guardXMLImplicitPrefixCase(xmlImplicitPrefixCaseID, selected, 52) == nil {
		t.Fatal("guard accepted implicit xml prefix ID or line drift")
	}
}

func implicitPrefixRoot(doc *losslessxml.Document) (losslessxml.Element, error) {
	if doc == nil {
		return losslessxml.Element{}, fmt.Errorf("missing implicit xml prefix document")
	}
	elements := doc.Elements()
	if len(elements) != 1 || elements[0].Name() != (xml.Name{Local: "r"}) {
		return losslessxml.Element{}, fmt.Errorf("implicit xml prefix root count or name drift")
	}
	root := elements[0]
	if _, ok := root.Parent(); ok || !bytes.Equal(root.Raw(), []byte(xmlImplicitPrefixInput)) {
		return losslessxml.Element{}, fmt.Errorf("implicit xml prefix root ownership/raw drift")
	}
	return root, nil
}

func registerXMLImplicitPrefixSteps(sc *godog.ScenarioContext, caller, original *[]byte, doc **losslessxml.Document) {
	sc.Step(`^the root namespace URI equals JSON ""$`, func() error {
		if !bytes.Equal(*original, []byte(xmlImplicitPrefixInput)) || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("implicit xml prefix root source drift")
		}
		_, err := implicitPrefixRoot(*doc)
		return err
	})
	sc.Step(`^the root attribute xml:lang equals JSON "en"$`, func() error {
		root, err := implicitPrefixRoot(*doc)
		if err != nil {
			return err
		}
		attrs := root.Attributes()
		if len(attrs) != 1 || attrs[0].Name != (xml.Name{Space: xmlImplicitPrefixURI, Local: "lang"}) || attrs[0].Value != "en" || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("implicit xml prefix expanded attribute/value or caller drift")
		}
		attrs[0].Value = "corrupt"
		again := root.Attributes()
		if len(again) != 1 || again[0].Name != (xml.Name{Space: xmlImplicitPrefixURI, Local: "lang"}) || again[0].Value != "en" {
			return fmt.Errorf("implicit xml prefix attributes alias snapshot")
		}
		return nil
	})
	sc.Step(`^the implicit xml namespace URI is (http://www\.w3\.org/XML/1998/namespace)$`, func(uri string) error {
		if uri != xmlImplicitPrefixURI {
			return fmt.Errorf("implicit xml prefix URI expectation drift")
		}
		root, err := implicitPrefixRoot(*doc)
		if err != nil {
			return err
		}
		ns := root.Namespaces()
		if len(ns) != 1 || ns["xml"] != uri || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("implicit xml prefix binding or caller drift")
		}
		ns["xml"] = "corrupt"
		if root.Namespaces()["xml"] != uri {
			return fmt.Errorf("implicit xml prefix namespace map aliases snapshot")
		}
		out, err := (*doc).Edit(nil, nil)
		if err != nil || !bytes.Equal(out, *original) || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("implicit xml prefix empty edit/caller drift: %v", err)
		}
		out[0] = '!'
		again, err := (*doc).Edit(nil, nil)
		if err != nil || !bytes.Equal(again, *original) || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("implicit xml prefix output aliases snapshot/caller: %v", err)
		}
		(*caller)[0] = '!'
		third, err := (*doc).Edit(nil, nil)
		if err != nil || !bytes.Equal(third, *original) {
			return fmt.Errorf("implicit xml prefix snapshot aliases caller: %v", err)
		}
		return nil
	})
}

func TestXMLImplicitPrefixRefusalAndRetry(t *testing.T) {
	bad := []byte(`<r xmlns:xml="urn:wrong" xml:lang="en"/>`)
	before := bytes.Clone(bad)
	invalid, err := losslessxml.Parse(bad)
	if err == nil || invalid != nil || !bytes.Equal(bad, before) {
		t.Fatalf("reserved xml rebind accepted or caller changed: %v", err)
	}
	good := []byte(xmlImplicitPrefixInput)
	original := bytes.Clone(good)
	doc, err := losslessxml.Parse(good)
	if err != nil || !bytes.Equal(good, original) {
		t.Fatalf("valid retry failed or changed caller: %v", err)
	}
	root, err := implicitPrefixRoot(doc)
	if err != nil {
		t.Fatal(err)
	}
	attrs := root.Attributes()
	if len(attrs) != 1 || attrs[0].Name != (xml.Name{Space: xmlImplicitPrefixURI, Local: "lang"}) || attrs[0].Value != "en" || root.Namespaces()["xml"] != xmlImplicitPrefixURI {
		t.Fatal("valid retry lost implicit xml prefix binding")
	}
}
