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

const elementRemovalSource = `<r xmlns:p="u"><!--keep--><p:a x = '1'><p:b>text</p:b></p:a> gap <p:c /></r>`
const elementRemovalResult = `<r xmlns:p="u"><!--keep--> gap </r>`

func guardElementRemovalCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"the XML source is " + elementRemovalSource,
		"the XML editor removes the p:a subtree and the p:c element from one parsed snapshot",
		"the complete output bytes equal " + elementRemovalResult,
		"a separate empty removal returns the exact original source bytes",
	}
	if id != elementRemovalCaseID || line != 72 || p.Name != "Remove two disjoint children without rewriting a comment or gap" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected element-removal canonical case %s %q at %d", id, p.Name, line)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("element-removal canonical step %d drift", i+1)
		}
	}
	return nil
}

func elementRemovalSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc = nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^the XML source is (<r xmlns:p="u"><!--keep--><p:a x = '1'><p:b>text</p:b></p:a> gap <p:c /></r>)$`, func(source string) error {
		if source != elementRemovalSource {
			return fmt.Errorf("element removal source drift")
		}
		caller = []byte(source)
		original = bytes.Clone(caller)
		return nil
	})
	sc.Step(`^the XML editor removes the p:a subtree and the p:c element from one parsed snapshot$`, func() error {
		// The shared Given step validates its literal; use a caller-owned copy
		// here so output and snapshot custody are checked independently.
		if !bytes.Equal(caller, []byte(elementRemovalSource)) {
			return fmt.Errorf("missing exact caller-owned removal source")
		}
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		es := doc.Elements()
		if len(es) != 4 || es[0].Name() != (xml.Name{Local: "r"}) || es[1].Name() != (xml.Name{Space: "u", Local: "a"}) || es[2].Name() != (xml.Name{Space: "u", Local: "b"}) || es[3].Name() != (xml.Name{Space: "u", Local: "c"}) {
			return fmt.Errorf("unexpected source-order removal targets")
		}
		if text, leaf := es[2].Text(); !leaf || text != "text" {
			return fmt.Errorf("nested text changed before removal")
		}
		output, err = doc.RemoveElements([]losslessxml.Element{es[1], es[3]})
		return err
	})
	sc.Step(`^the complete output bytes equal (<r xmlns:p="u"><!--keep--> gap </r>)$`, func(expected string) error {
		if doc == nil || expected != elementRemovalResult || !bytes.Equal(output, []byte(elementRemovalResult)) || !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(elementRemovalSource)) {
			return fmt.Errorf("removal output/caller bytes differ: %q", output)
		}
		prefix := []byte(`<r xmlns:p="u"><!--keep-->`)
		suffix := []byte(` gap </r>`)
		if !bytes.HasPrefix(output, prefix) || !bytes.HasSuffix(output, suffix) || bytes.Count(output, []byte(`<!--keep-->`)) != 1 {
			return fmt.Errorf("comment or gap changed")
		}
		reparsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		if es := reparsed.Elements(); len(es) != 1 || es[0].Name() != (xml.Name{Local: "r"}) {
			return fmt.Errorf("removed element survived")
		}
		return nil
	})
	sc.Step(`^a separate empty removal returns the exact original source bytes$`, func() error {
		if doc == nil {
			return fmt.Errorf("missing parsed snapshot")
		}
		empty, err := doc.RemoveElements(nil)
		if err != nil || !bytes.Equal(empty, []byte(elementRemovalSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("empty removal changed snapshot or caller: %v", err)
		}
		es := doc.Elements()
		for _, targets := range [][]losslessxml.Element{{es[0]}, {es[1], es[2]}} {
			refused, refusal := doc.RemoveElements(targets)
			if refusal == nil || refused != nil {
				return fmt.Errorf("root or overlapping removal accepted/output exposed")
			}
		}
		again, err := doc.RemoveElements([]losslessxml.Element{es[1], es[3]})
		if err != nil || !bytes.Equal(again, []byte(elementRemovalResult)) {
			return fmt.Errorf("refusal consumed snapshot: %v", err)
		}
		output[0] = '!'
		retained, err := doc.RemoveElements(nil)
		if err != nil || !bytes.Equal(retained, []byte(elementRemovalSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("returned output aliases snapshot/caller: %v", err)
		}
		return nil
	})
}
