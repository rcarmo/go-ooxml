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

const childInsertionSource = `<root xmlns="u" xmlns:n1="occupied"><a/><b>keep</b></root>`

func guardChildInsertionCustodyCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"the XML source is " + childInsertionSource,
		"a structured edit inserts a new-namespace x child with an other-namespace a attribute and a plain text child under the existing a element",
		"reparsing finds expanded element names new/x and empty-namespace plain",
		"the unedited sibling bytes <b>keep</b> remain in the output",
	}
	if contract20Candidate() {
		steps = append(steps[:3], append([]string{
			`the authored attribute expanded name is other/a with JSON value "value" and the plain grandchild text is JSON "text"`,
		}, steps[3:]...)...)
		steps = append(steps, uniformHistoricalAssertion)
	}
	if id != childInsertionCustodyCaseID || line != historicalEditingLine(38, 47, 49) || p.Name != "Insert a child with independently scoped element and attribute names" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected child-insertion canonical case %s %q at %d", id, p.Name, line)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step || (contract20Candidate() && p.Steps[i].Argument != nil) {
			return fmt.Errorf("child-insertion canonical step %d drift", i+1)
		}
	}
	return nil
}

func childInsertionNode() losslessxml.NewElement {
	return losslessxml.NewElement{
		Name:       xml.Name{Space: "new", Local: "x"},
		Attributes: []xml.Attr{{Name: xml.Name{Space: "other", Local: "a"}, Value: childInsertionAttributeValue()}},
		Children:   []losslessxml.NewElement{{Name: xml.Name{Local: "plain"}, Text: "text"}},
	}
}

func childInsertionAttributeValue() string {
	if contract20Candidate() {
		return "value"
	}
	return "v"
}

func checkChildInsertionResult(output []byte) error {
	parsed, err := losslessxml.Parse(output)
	if err != nil {
		return err
	}
	es := parsed.Elements()
	want := []xml.Name{{Space: "u", Local: "root"}, {Space: "u", Local: "a"}, {Space: "new", Local: "x"}, {Local: "plain"}, {Space: "u", Local: "b"}}
	if len(es) != len(want) {
		return fmt.Errorf("inserted element count %d", len(es))
	}
	for i, name := range want {
		if es[i].Name() != name {
			return fmt.Errorf("inserted element %d name %v", i, es[i].Name())
		}
	}
	for i, parent := range []int{0, 1, 2, 0} {
		p, ok := es[i+1].Parent()
		if !ok || p != es[parent] {
			return fmt.Errorf("inserted element %d parent drift", i+1)
		}
	}
	attrs := es[2].Attributes()
	if len(attrs) != 1 || attrs[0].Name != (xml.Name{Space: "other", Local: "a"}) || attrs[0].Value != childInsertionAttributeValue() {
		return fmt.Errorf("inserted attribute drift: %+v", attrs)
	}
	if text, leaf := es[3].Text(); !leaf || text != "text" {
		return fmt.Errorf("plain child text drift")
	}
	if text, leaf := es[4].Text(); !leaf || text != "keep" {
		return fmt.Errorf("untouched sibling text drift")
	}
	return nil
}

func childInsertionCustodySteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var parent losslessxml.Element
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, parent = nil, nil, nil, nil, losslessxml.Element{}
		return ctx, nil
	})
	sc.Step(`^the XML source is (<root xmlns="u" xmlns:n1="occupied"><a/><b>keep</b></root>)$`, func(source string) error {
		if source != childInsertionSource {
			return fmt.Errorf("child-insertion source drift")
		}
		caller = []byte(source)
		original = bytes.Clone(caller)
		return nil
	})
	sc.Step(`^a structured edit inserts a new-namespace x child with an other-namespace a attribute and a plain text child under the existing a element$`, func() error {
		if !bytes.Equal(caller, []byte(childInsertionSource)) {
			return fmt.Errorf("missing caller-owned insertion source")
		}
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		es := doc.Elements()
		if len(es) != 3 || es[0].Name() != (xml.Name{Space: "u", Local: "root"}) || es[1].Name() != (xml.Name{Space: "u", Local: "a"}) || es[2].Name() != (xml.Name{Space: "u", Local: "b"}) {
			return fmt.Errorf("source-order insertion target drift")
		}
		if text, leaf := es[2].Text(); !leaf || text != "keep" {
			return fmt.Errorf("original sibling text drift")
		}
		parent = es[1]
		empty, err := doc.InsertChildren(nil)
		if err != nil || !bytes.Equal(empty, original) {
			return fmt.Errorf("empty insertion changed original: %v", err)
		}
		output, err = doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: parent, Children: []losslessxml.NewElement{childInsertionNode()}}})
		return err
	})
	sc.Step(`^reparsing finds expanded element names new/x and empty-namespace plain$`, func() error {
		if doc == nil || !bytes.Equal(caller, original) {
			return fmt.Errorf("insertion changed caller input")
		}
		return checkChildInsertionResult(output)
	})
	sc.Step(`^the unedited sibling bytes <b>keep</b> remain in the output$`, func() error {
		prefix := []byte(`<root xmlns="u" xmlns:n1="occupied"><a>`)
		suffix := []byte(`</a><b>keep</b></root>`)
		if !bytes.HasPrefix(output, prefix) || !bytes.HasSuffix(output, suffix) || bytes.Count(output, []byte(`<b>keep</b>`)) != 1 {
			return fmt.Errorf("sibling or surviving source bytes changed: %q", output)
		}
		if !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(childInsertionSource)) {
			return fmt.Errorf("caller bytes changed after insertion")
		}
		again, err := doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: parent, Children: []losslessxml.NewElement{childInsertionNode()}}})
		if err != nil || !bytes.Equal(again, output) {
			return fmt.Errorf("parsed snapshot consumed after insertion: %v", err)
		}
		output[0] = '!'
		empty, err := doc.InsertChildren(nil)
		if err != nil || !bytes.Equal(empty, []byte(childInsertionSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("returned bytes alias snapshot or caller: %v", err)
		}
		return checkChildInsertionResult(again)
	})
}
