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

func guardChildInsertionRefusalCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"the XML source is " + childInsertionSource,
		"one structured insertion batch targets both the root and its nested a element",
		"the insertion returns an error",
		"a separate empty insertion batch returns the exact original source bytes",
	}
	if id != childInsertionRefusalCaseID || line != lexicalEditingLine(45, 54) || p.Name != "Overlapping insertion targets refuse and a separate no-op preserves bytes" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected child-insertion refusal case %s %q at %d", id, p.Name, line)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("child-insertion refusal step %d drift", i+1)
		}
	}
	return nil
}

func childInsertionRefusalSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var root, child losslessxml.Element
	var refusal error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, root, child, refusal = nil, nil, nil, nil, losslessxml.Element{}, losslessxml.Element{}, nil
		return ctx, nil
	})
	sc.Step(`^one structured insertion batch targets both the root and its nested a element$`, func() error {
		// The shared Given validates the source. Keep a separate caller-owned
		// slice for this refusal's output and snapshot custody assertions.
		caller = []byte(childInsertionSource)
		original = bytes.Clone(caller)
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		es := doc.Elements()
		if len(es) != 3 || es[0].Name() != (xml.Name{Space: "u", Local: "root"}) || es[1].Name() != (xml.Name{Space: "u", Local: "a"}) || es[2].Name() != (xml.Name{Space: "u", Local: "b"}) {
			return fmt.Errorf("source-order insertion targets drift")
		}
		p, ok := es[1].Parent()
		if !ok || p != es[0] {
			return fmt.Errorf("a is not nested under root")
		}
		if text, leaf := es[2].Text(); !leaf || text != "keep" {
			return fmt.Errorf("source sibling text drift")
		}
		root, child = es[0], es[1]
		output, refusal = doc.InsertChildren([]losslessxml.ChildInsertion{
			{Parent: root, Children: []losslessxml.NewElement{{Name: xml.Name{Local: "atRoot"}, Text: "root text"}}},
			{Parent: child, Children: []losslessxml.NewElement{childInsertionNode()}},
		})
		return nil // Then checks both the error and absence of output.
	})
	sc.Step(`^the insertion returns an error$`, func() error {
		if doc == nil || refusal == nil || output != nil {
			return fmt.Errorf("nested insertion yielded output %q, error %v", output, refusal)
		}
		if !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(childInsertionSource)) {
			return fmt.Errorf("refusal changed caller bytes")
		}
		es := doc.Elements()
		if len(es) != 3 || es[0] != root || es[1] != child || es[2].Name() != (xml.Name{Space: "u", Local: "b"}) {
			return fmt.Errorf("refusal changed parsed source targets")
		}
		if text, leaf := es[2].Text(); !leaf || text != "keep" {
			return fmt.Errorf("refusal changed snapshot sibling text")
		}
		return nil
	})
	sc.Step(`^a separate empty insertion batch returns the exact original source bytes$`, func() error {
		if doc == nil || refusal == nil || output != nil {
			return fmt.Errorf("missing preceding nested-target refusal")
		}
		empty, err := doc.InsertChildren(nil)
		if err != nil || !bytes.Equal(empty, []byte(childInsertionSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("empty insertion after refusal changed source: %v", err)
		}
		valid, err := doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: child, Children: []losslessxml.NewElement{childInsertionNode()}}})
		if err != nil || !bytes.HasPrefix(valid, []byte(`<root xmlns="u" xmlns:n1="occupied"><a>`)) || !bytes.HasSuffix(valid, []byte(`</a><b>keep</b></root>`)) || bytes.Count(valid, []byte(`<b>keep</b>`)) != 1 {
			return fmt.Errorf("valid a-only insertion after refusal lost source bytes: %q %v", valid, err)
		}
		if err := checkChildInsertionResult(valid); err != nil {
			return fmt.Errorf("valid a-only insertion after refusal: %w", err)
		}
		valid[0] = '!'
		again, err := doc.InsertChildren(nil)
		if err != nil || !bytes.Equal(again, []byte(childInsertionSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("returned output aliases caller or snapshot: %v", err)
		}
		return nil
	})
}
