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

const attributeRefusalEdited = `<r xmlns:p="urn:p"><t a = 'x' p:n="old"/><t>text</t></r>`

func guardAttributeRefusalCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"the XML source is " + attributeSpliceSource,
		"one lexical edit batch sets a of the first t element to x and to y",
		"the edit returns an error instead of accepting that batch",
	}
	if id != attributeRefusalCaseID || line != historicalEditingLine(32, 41, 42) || p.Name != "Duplicate edits to one attribute refuse the batch" || len(p.AstNodeIds) != 1 || !historicalStepCount(p, len(steps)) {
		return fmt.Errorf("unexpected attribute-refusal canonical case %s %q at %d", id, p.Name, line)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("attribute-refusal canonical step %d drift", i+1)
		}
	}
	return nil
}

func attributeRefusalSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var refusal error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, refusal = nil, nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^one lexical edit batch sets a of the first t element to x and to y$`, func() error {
		// The shared Given binding validates its source text; keep an independent
		// caller-owned slice for this canonical refusal's custody assertions.
		caller = []byte(attributeSpliceSource)
		original = bytes.Clone(caller)
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		if err = checkAttributeSpliceDocument(doc, [3]string{"a&b", "old", ""}); err != nil {
			return fmt.Errorf("input: %w", err)
		}
		first := doc.Elements()[1]
		attribute := xml.Name{Local: "a"}
		output, refusal = doc.Edit(nil, []losslessxml.AttributeEdit{
			{Target: first, Name: attribute, Value: "x"},
			{Target: first, Name: attribute, Value: "y"},
		})
		return nil // The Then step independently checks both refusal and output custody.
	})
	sc.Step(`^the edit returns an error instead of accepting that batch$`, func() error {
		if doc == nil || refusal == nil || output != nil {
			return fmt.Errorf("duplicate-a batch returned output %q, error %v", output, refusal)
		}
		if !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(attributeSpliceSource)) {
			return fmt.Errorf("duplicate-a batch changed caller input")
		}
		if err := checkAttributeSpliceDocument(doc, [3]string{"a&b", "old", ""}); err != nil {
			return fmt.Errorf("duplicate-a batch changed parsed snapshot: %w", err)
		}
		empty, err := doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(empty, []byte(attributeSpliceSource)) {
			return fmt.Errorf("empty edit after refusal differs: %v", err)
		}
		same, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "a&b"}})
		if err != nil || !bytes.Equal(same, []byte(attributeSpliceSource)) {
			return fmt.Errorf("same-value edit after refusal differs: %v", err)
		}
		// A valid edit on the very same snapshot proves refusal did not consume it.
		changed, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "x"}})
		if err != nil || !bytes.Equal(changed, []byte(attributeRefusalEdited)) {
			return fmt.Errorf("valid edit after refusal bytes %q: %v", changed, err)
		}
		reparsed, err := losslessxml.Parse(changed)
		if err != nil {
			return err
		}
		if err := checkAttributeSpliceDocument(reparsed, [3]string{"x", "old", ""}); err != nil {
			return fmt.Errorf("valid edit after refusal: %w", err)
		}
		// Reversed values are a separate auxiliary control, not another case.
		reversed, reverseErr := doc.Edit(nil, []losslessxml.AttributeEdit{
			{Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "y"},
			{Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "x"},
		})
		if reverseErr == nil || reversed != nil {
			return fmt.Errorf("reversed duplicate-a accepted or emitted bytes")
		}
		changed[0] = '!'
		retained, err := doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(retained, []byte(attributeSpliceSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("returned bytes alias caller or snapshot: %v", err)
		}
		return nil
	})
}
