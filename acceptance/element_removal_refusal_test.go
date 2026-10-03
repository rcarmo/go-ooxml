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

func guardElementRemovalRefusalCase(id string, p *messages.Pickle, line int) (string, error) {
	base := historicalEditingLine(85, 94, 101)
	rows := map[int]string{base: "root", base + 1: "p:a and its nested p:b"}
	selection, ok := rows[line]
	if id != elementRemovalRefusalCaseID || !ok || len(p.AstNodeIds) != 2 || p.Name != "A "+selection+" removal refuses" || !historicalStepCount(p, 3) {
		return "", fmt.Errorf("unexpected element-removal refusal case %s %q at %d", id, p.Name, line)
	}
	steps := []string{
		"the XML source is " + elementRemovalSource,
		"a removal batch selects " + selection,
		"the removal returns an error",
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return "", fmt.Errorf("element-removal refusal step %d drift for %s", i+1, selection)
		}
	}
	return selection, nil
}

func elementRemovalRefusalSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var refusal error
	var selection string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, refusal, selection = nil, nil, nil, nil, nil, ""
		return ctx, nil
	})
	sc.Step(`^a removal batch selects (root|p:a and its nested p:b)$`, func(chosen string) error {
		// The shared Given checks its literal. Recreate caller-owned input here;
		// the result, caller, and immutable snapshot are checked separately.
		caller = []byte(elementRemovalSource)
		original = bytes.Clone(caller)
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		es := doc.Elements()
		if len(es) != 4 || es[0].Name() != (xml.Name{Local: "r"}) || es[1].Name() != (xml.Name{Space: "u", Local: "a"}) || es[2].Name() != (xml.Name{Space: "u", Local: "b"}) || es[3].Name() != (xml.Name{Space: "u", Local: "c"}) {
			return fmt.Errorf("unexpected source-order refusal targets")
		}
		selection = chosen
		targets := []losslessxml.Element{es[0]}
		if chosen == "p:a and its nested p:b" {
			targets = []losslessxml.Element{es[1], es[2]}
		}
		output, refusal = doc.RemoveElements(targets)
		return nil // Then checks both refusal and nil output independently.
	})
	sc.Step(`^the removal returns an error$`, func() error {
		if doc == nil || selection == "" || refusal == nil || output != nil {
			return fmt.Errorf("%s removal yielded output %q, error %v", selection, output, refusal)
		}
		if !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(elementRemovalSource)) {
			return fmt.Errorf("%s refusal changed caller bytes", selection)
		}
		es := doc.Elements()
		if len(es) != 4 || es[0].Name() != (xml.Name{Local: "r"}) || es[1].Name() != (xml.Name{Space: "u", Local: "a"}) || es[2].Name() != (xml.Name{Space: "u", Local: "b"}) || es[3].Name() != (xml.Name{Space: "u", Local: "c"}) {
			return fmt.Errorf("%s refusal changed the parsed snapshot", selection)
		}
		if text, leaf := es[2].Text(); !leaf || text != "text" {
			return fmt.Errorf("%s refusal changed nested text", selection)
		}
		empty, err := doc.RemoveElements(nil)
		if err != nil || !bytes.Equal(empty, []byte(elementRemovalSource)) {
			return fmt.Errorf("%s empty removal after refusal changed bytes: %v", selection, err)
		}
		valid, err := doc.RemoveElements([]losslessxml.Element{es[1], es[3]})
		if err != nil || !bytes.Equal(valid, []byte(elementRemovalResult)) {
			return fmt.Errorf("%s refusal consumed snapshot or changed valid removal: %q %v", selection, valid, err)
		}
		if !bytes.Equal(caller, original) {
			return fmt.Errorf("%s valid retry changed caller bytes", selection)
		}
		again, err := doc.RemoveElements(nil)
		if err != nil || !bytes.Equal(again, []byte(elementRemovalSource)) {
			return fmt.Errorf("%s valid retry consumed snapshot: %v", selection, err)
		}
		return nil
	})
}
