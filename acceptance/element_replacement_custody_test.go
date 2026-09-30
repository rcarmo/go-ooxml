package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const elementReplacementCustodyCaseID = "@id-xml-go-element-replacement-custody"
const elementReplacementSource = `<root xmlns="outer" xmlns:p="bound"><!--a--><p:old xmlns:p="inner" x='1'><p:child/></p:old> tail <last/></root>`
const elementReplacementOutput = `<root xmlns="outer" xmlns:p="bound"><!--a--><p:new>value</p:new><plain xmlns=""/> tail <last/></root>`

func guardElementReplacementCustodyCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"the XML source is " + elementReplacementSource,
		"the XML editor replaces p:old with a bound-namespace new element containing value and an empty-namespace plain element",
		"the complete output bytes equal " + elementReplacementOutput,
	}
	if id != elementReplacementCustodyCaseID || line != lexicalEditingLine(89, 98) || p.Name != "Replace one subtree using its surviving parent namespace scope" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected element-replacement custody case %s %q at %d", id, p.Name, line)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("element-replacement custody step %d drift", i+1)
		}
	}
	return nil
}

func TestElementReplacementCustodyGuardRejectsDrift(t *testing.T) {
	path := xmlEditingFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil {
		t.Fatal(err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == elementReplacementCustodyCaseID {
				if selected != nil {
					t.Fatal("duplicate element-replacement custody case")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardElementReplacementCustodyCase(elementReplacementCustodyCaseID, selected, lexicalEditingLine(89, 98)) != nil {
		t.Fatal("canonical element-replacement custody guard failed")
	}
	for _, tc := range []struct {
		name   string
		mutate func(*messages.Pickle)
	}{
		{"source", func(p *messages.Pickle) { p.Steps[0].Text += " " }},
		{"whole output", func(p *messages.Pickle) { p.Steps[2].Text += " " }},
		{"action", func(p *messages.Pickle) { p.Steps[1].Text += " changed" }},
		{"scenario name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"example expansion", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "unexpected") }},
		{"step argument", func(p *messages.Pickle) { p.Steps[0].Argument = &messages.PickleStepArgument{} }},
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
			if guardElementReplacementCustodyCase(elementReplacementCustodyCaseID, &clone, lexicalEditingLine(89, 98)) == nil {
				t.Fatal("guard accepted canonical drift")
			}
		})
	}
	if guardElementReplacementCustodyCase("@id-xml-go-element-replacement-refusal", selected, lexicalEditingLine(89, 98)) == nil || guardElementReplacementCustodyCase(elementReplacementCustodyCaseID, selected, lexicalEditingLine(89, 98)+1) == nil {
		t.Fatal("guard accepted ID or line drift")
	}
}

func elementReplacementNodes() []losslessxml.NewElement {
	return []losslessxml.NewElement{
		{Name: xml.Name{Space: "bound", Local: "new"}, Text: "value"},
		{Name: xml.Name{Local: "plain"}},
	}
}

func elementReplacementCustodySteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var target losslessxml.Element
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, target = nil, nil, nil, nil, losslessxml.Element{}
		return ctx, nil
	})
	sc.Step(`^the XML source is (<root xmlns="outer" xmlns:p="bound"><!--a--><p:old xmlns:p="inner" x='1'><p:child/></p:old> tail <last/></root>)$`, func(source string) error {
		if source != elementReplacementSource {
			return fmt.Errorf("element-replacement source drift")
		}
		caller = []byte(source)
		original = bytes.Clone(caller)
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		es := doc.Elements()
		if len(es) != 4 || es[0].Name() != (xml.Name{Space: "outer", Local: "root"}) || es[1].Name() != (xml.Name{Space: "inner", Local: "old"}) || es[2].Name() != (xml.Name{Space: "inner", Local: "child"}) || es[3].Name() != (xml.Name{Space: "outer", Local: "last"}) {
			return fmt.Errorf("element-replacement original expanded-name or source-order drift")
		}
		if parent, ok := es[1].Parent(); !ok || parent != es[0] {
			return fmt.Errorf("element-replacement original target parent drift")
		}
		if parent, ok := es[2].Parent(); !ok || parent != es[1] {
			return fmt.Errorf("element-replacement original child parent drift")
		}
		if parent, ok := es[3].Parent(); !ok || parent != es[0] {
			return fmt.Errorf("element-replacement original last parent drift")
		}
		if attrs := es[1].Attributes(); len(attrs) != 1 || attrs[0].Name != (xml.Name{Local: "x"}) || attrs[0].Value != "1" {
			return fmt.Errorf("element-replacement original target attribute drift")
		}
		if ns := es[0].Namespaces(); ns[""] != "outer" || ns["p"] != "bound" {
			return fmt.Errorf("element-replacement surviving parent namespace drift")
		}
		if ns := es[1].Namespaces(); ns["p"] != "inner" {
			return fmt.Errorf("element-replacement removed target namespace drift")
		}
		target = es[1]
		return nil
	})
	sc.Step(`^the XML editor replaces p:old with a bound-namespace new element containing value and an empty-namespace plain element$`, func() error {
		if doc == nil || !bytes.Equal(caller, original) {
			return fmt.Errorf("missing or changed caller-owned replacement source")
		}
		var err error
		output, err = doc.ReplaceElements([]losslessxml.ElementReplacement{{Target: target, Nodes: elementReplacementNodes()}})
		if err != nil || !bytes.Equal(caller, original) {
			return fmt.Errorf("replacement failed or changed caller: %v", err)
		}
		return nil
	})
	sc.Step(`^the complete output bytes equal (<root xmlns="outer" xmlns:p="bound"><!--a--><p:new>value</p:new><plain xmlns=""/> tail <last/></root>)$`, func(expected string) error {
		if expected != elementReplacementOutput || !bytes.Equal(output, []byte(elementReplacementOutput)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("element-replacement whole output or caller drift: %q", output)
		}
		parsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		es := parsed.Elements()
		want := []xml.Name{{Space: "outer", Local: "root"}, {Space: "bound", Local: "new"}, {Local: "plain"}, {Space: "outer", Local: "last"}}
		if len(es) != len(want) {
			return fmt.Errorf("element-replacement output element count %d", len(es))
		}
		for i, name := range want {
			if es[i].Name() != name {
				return fmt.Errorf("element-replacement output element %d name %v", i, es[i].Name())
			}
			if i > 0 {
				parent, ok := es[i].Parent()
				if !ok || parent != es[0] {
					return fmt.Errorf("element-replacement output element %d parent drift", i)
				}
			}
		}
		if text, leaf := es[1].Text(); !leaf || text != "value" {
			return fmt.Errorf("element-replacement new text drift")
		}
		if text, leaf := es[2].Text(); !leaf || text != "" {
			return fmt.Errorf("element-replacement plain text drift")
		}
		if attrs := es[1].Attributes(); len(attrs) != 0 {
			return fmt.Errorf("element-replacement inherited removed attribute")
		}
		if !bytes.HasPrefix(output, []byte(`<root xmlns="outer" xmlns:p="bound"><!--a-->`)) || !bytes.HasSuffix(output, []byte(` tail <last/></root>`)) {
			return fmt.Errorf("element-replacement comment/gap/last drift")
		}
		empty, err := doc.ReplaceElements(nil)
		if err != nil || !bytes.Equal(empty, original) || !bytes.Equal(caller, original) {
			return fmt.Errorf("element-replacement snapshot or caller changed: %v", err)
		}
		repeated, err := doc.ReplaceElements([]losslessxml.ElementReplacement{{Target: target, Nodes: elementReplacementNodes()}})
		if err != nil || !bytes.Equal(repeated, []byte(elementReplacementOutput)) {
			return fmt.Errorf("element-replacement snapshot consumed: %v", err)
		}
		output[0] = '!'
		empty, err = doc.ReplaceElements(nil)
		if err != nil || !bytes.Equal(empty, original) || !bytes.Equal(caller, original) || !bytes.Equal(repeated, []byte(elementReplacementOutput)) {
			return fmt.Errorf("element-replacement returned bytes alias snapshot or caller: %v", err)
		}
		return nil
	})
}
