package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const elementReplacementRefusalCaseID = "@id-xml-go-element-replacement-refusal"

func elementReplacementRefusalRows() map[int]string {
	base := lexicalEditingLine(102, 111)
	return map[int]string{base: "root", base + 1: "p:old twice", base + 2: "p:old and its nested p:child"}
}

func guardElementReplacementRefusalCase(id string, p *messages.Pickle, line int) (string, error) {
	selection, ok := elementReplacementRefusalRows()[line]
	if id != elementReplacementRefusalCaseID || !ok || len(p.AstNodeIds) != 2 || p.Name != "A "+selection+" subtree replacement refuses without output" || len(p.Steps) != 4 {
		return "", fmt.Errorf("unexpected element-replacement refusal case %s %q at %d", id, p.Name, line)
	}
	steps := []string{
		"the XML source is " + elementReplacementSource,
		"a replacement batch selects " + selection,
		"replacement returns an error and no edited output",
		"a separate empty replacement returns the exact original source bytes",
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return "", fmt.Errorf("element-replacement refusal step %d drift for %s", i+1, selection)
		}
	}
	return selection, nil
}

func checkReplacementRefusalSnapshot(doc *losslessxml.Document) error {
	if doc == nil {
		return fmt.Errorf("missing element-replacement snapshot")
	}
	es := doc.Elements()
	want := []xml.Name{{Space: "outer", Local: "root"}, {Space: "inner", Local: "old"}, {Space: "inner", Local: "child"}, {Space: "outer", Local: "last"}}
	if len(es) != len(want) {
		return fmt.Errorf("element-replacement refusal original element count %d", len(es))
	}
	for i, name := range want {
		if es[i].Name() != name {
			return fmt.Errorf("element-replacement refusal original element %d name %v", i, es[i].Name())
		}
		if i > 0 {
			parentIndex := 0
			if i == 2 {
				parentIndex = 1
			}
			parent, ok := es[i].Parent()
			if !ok || parent != es[parentIndex] {
				return fmt.Errorf("element-replacement refusal original element %d parent drift", i)
			}
		}
	}
	attrs := es[1].Attributes()
	if len(attrs) != 1 || attrs[0].Name != (xml.Name{Local: "x"}) || attrs[0].Value != "1" {
		return fmt.Errorf("element-replacement refusal original attribute drift")
	}
	if ns := es[0].Namespaces(); ns[""] != "outer" || ns["p"] != "bound" {
		return fmt.Errorf("element-replacement refusal original parent namespace drift")
	}
	if ns := es[1].Namespaces(); ns["p"] != "inner" {
		return fmt.Errorf("element-replacement refusal original target namespace drift")
	}
	return nil
}

func elementReplacementRefusalSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var refusal error
	var selection string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, refusal, selection = nil, nil, nil, nil, nil, ""
		return ctx, nil
	})
	// The custody Given validates the exact canonical source. Build a separate
	// caller-owned snapshot for each outline row in this scenario's When step.
	sc.Step(`^a replacement batch selects (root|p:old twice|p:old and its nested p:child)$`, func(chosen string) error {
		if _, ok := map[string]bool{"root": true, "p:old twice": true, "p:old and its nested p:child": true}[chosen]; !ok {
			return fmt.Errorf("unexpected replacement refusal selection %q", chosen)
		}
		caller = []byte(elementReplacementSource)
		original = bytes.Clone(caller)
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		if err := checkReplacementRefusalSnapshot(doc); err != nil {
			return err
		}
		es := doc.Elements()
		replacement := func(target losslessxml.Element) losslessxml.ElementReplacement {
			return losslessxml.ElementReplacement{Target: target, Nodes: elementReplacementNodes()}
		}
		var edits []losslessxml.ElementReplacement
		switch chosen {
		case "root":
			edits = []losslessxml.ElementReplacement{replacement(es[0])}
		case "p:old twice":
			edits = []losslessxml.ElementReplacement{replacement(es[1]), replacement(es[1])}
		case "p:old and its nested p:child":
			edits = []losslessxml.ElementReplacement{replacement(es[1]), replacement(es[2])}
		}
		selection = chosen
		output, refusal = doc.ReplaceElements(edits)
		return nil // Then checks the error and nil output separately.
	})
	sc.Step(`^replacement returns an error and no edited output$`, func() error {
		if doc == nil || selection == "" || refusal == nil || output != nil {
			return fmt.Errorf("%s replacement yielded output %q, error %v", selection, output, refusal)
		}
		if !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(elementReplacementSource)) {
			return fmt.Errorf("%s refusal changed caller bytes", selection)
		}
		return checkReplacementRefusalSnapshot(doc)
	})
	sc.Step(`^a separate empty replacement returns the exact original source bytes$`, func() error {
		if doc == nil || selection == "" || refusal == nil || output != nil {
			return fmt.Errorf("missing %s refusal before empty replacement", selection)
		}
		empty, err := doc.ReplaceElements(nil)
		if err != nil || !bytes.Equal(empty, original) || !bytes.Equal(caller, original) {
			return fmt.Errorf("%s empty replacement changed source: %v", selection, err)
		}
		if err := checkReplacementRefusalSnapshot(doc); err != nil {
			return err
		}
		es := doc.Elements()
		valid, err := doc.ReplaceElements([]losslessxml.ElementReplacement{{Target: es[1], Nodes: elementReplacementNodes()}})
		if err != nil || !bytes.Equal(valid, []byte(elementReplacementOutput)) {
			return fmt.Errorf("%s valid retry failed: %q %v", selection, valid, err)
		}
		parsed, err := losslessxml.Parse(valid)
		if err != nil {
			return err
		}
		want := []xml.Name{{Space: "outer", Local: "root"}, {Space: "bound", Local: "new"}, {Local: "plain"}, {Space: "outer", Local: "last"}}
		got := parsed.Elements()
		if len(got) != len(want) {
			return fmt.Errorf("%s valid retry element count drift", selection)
		}
		for i, name := range want {
			if got[i].Name() != name {
				return fmt.Errorf("%s valid retry element %d name drift", selection, i)
			}
			if i > 0 {
				parent, ok := got[i].Parent()
				if !ok || parent != got[0] {
					return fmt.Errorf("%s valid retry element %d parent drift", selection, i)
				}
			}
		}
		if text, leaf := got[1].Text(); !leaf || text != "value" {
			return fmt.Errorf("%s valid retry text drift", selection)
		}
		if text, leaf := got[2].Text(); !leaf || text != "" {
			return fmt.Errorf("%s valid retry plain text drift", selection)
		}
		if attrs := got[1].Attributes(); len(attrs) != 0 {
			return fmt.Errorf("%s valid retry retained old attributes", selection)
		}
		if !bytes.Equal(caller, original) {
			return fmt.Errorf("%s valid retry changed caller", selection)
		}
		valid[0] = '!'
		empty, err = doc.ReplaceElements(nil)
		if err != nil || !bytes.Equal(empty, original) || !bytes.Equal(caller, original) {
			return fmt.Errorf("%s valid result aliases snapshot or caller: %v", selection, err)
		}
		again, err := doc.ReplaceElements([]losslessxml.ElementReplacement{{Target: es[1], Nodes: elementReplacementNodes()}})
		if err != nil || !bytes.Equal(again, []byte(elementReplacementOutput)) || checkReplacementRefusalSnapshot(doc) != nil {
			return fmt.Errorf("%s valid retry consumed snapshot: %v", selection, err)
		}
		return nil
	})
}
