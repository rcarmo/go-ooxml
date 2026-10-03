package acceptance

import (
	"bytes"
	"context"
	"fmt"

	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const immutableLeafSource = `<r><t>hello</t></r>`
const immutableLeafEdited = `<r><t>bye &amp; &lt;</t></r>`

func guardImmutableLeafCase(id string, p *messages.Pickle) error {
	steps := []string{
		"the XML source is <r><t>hello</t></r>",
		"the XML editor parses a caller-owned byte slice and performs an empty edit",
		"the caller input bytes still equal the original XML source",
		"the empty edit returns the exact original source bytes",
	}
	if id != immutableLeafCaseID || p.Name != "A seeded immutable parse and no-op leave the caller bytes and parsed snapshot intact" || !historicalStepCount(p, len(steps)) {
		return fmt.Errorf("unexpected immutable-leaf canonical case %s %q", id, p.Name)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("immutable-leaf canonical step %d drift", i+1)
		}
	}
	return nil
}

func immutableLeafSteps(sc *godog.ScenarioContext) {
	var caller, before, output []byte
	var doc *losslessxml.Document
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, before, output, doc = nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^the XML source is <r><t>hello</t></r>$`, func() error {
		caller = []byte(immutableLeafSource)
		before = bytes.Clone(caller)
		return nil
	})
	sc.Step(`^the XML editor parses a caller-owned byte slice and performs an empty edit$`, func() error {
		if !bytes.Equal(caller, []byte(immutableLeafSource)) {
			return fmt.Errorf("wrong immutable-leaf source")
		}
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		elements := doc.Elements()
		if len(elements) != 2 || elements[0].Name().Local != "r" || elements[0].Name().Space != "" || elements[1].Name().Local != "t" || elements[1].Name().Space != "" {
			return fmt.Errorf("unexpected root or leaf")
		}
		if text, leaf := elements[1].Text(); !leaf || text != "hello" {
			return fmt.Errorf("unexpected leaf text %q leaf=%v", text, leaf)
		}
		output, err = doc.Edit(nil, nil)
		return err
	})
	sc.Step(`^the caller input bytes still equal the original XML source$`, func() error {
		if doc == nil || !bytes.Equal(caller, before) || !bytes.Equal(caller, []byte(immutableLeafSource)) {
			return fmt.Errorf("caller input changed")
		}
		return nil
	})
	sc.Step(`^the empty edit returns the exact original source bytes$`, func() error {
		if doc == nil || !bytes.Equal(output, []byte(immutableLeafSource)) || !bytes.Equal(output, before) {
			return fmt.Errorf("empty edit changed original bytes: %q", output)
		}
		// A nonempty edit on this same parsed leaf must not be a pass-through.
		leaf := doc.Elements()[1]
		changed, err := doc.Edit([]losslessxml.TextEdit{{Target: leaf, Text: "bye & <"}}, nil)
		if err != nil || !bytes.Equal(changed, []byte(immutableLeafEdited)) {
			return fmt.Errorf("nonempty edit bytes %q: %v", changed, err)
		}
		reparsed, err := losslessxml.Parse(changed)
		if err != nil {
			return err
		}
		if elems := reparsed.Elements(); len(elems) != 2 {
			return fmt.Errorf("edited leaf missing")
		} else if text, ok := elems[1].Text(); !ok || text != "bye & <" {
			return fmt.Errorf("edited text %q leaf=%v", text, ok)
		}
		if !bytes.Equal(caller, before) {
			return fmt.Errorf("edit mutated caller input")
		}
		// The returned no-op bytes must not alias either the input or the snapshot.
		output[0] = '!'
		unchanged, err := doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(unchanged, []byte(immutableLeafSource)) || !bytes.Equal(caller, before) {
			return fmt.Errorf("returned bytes alias caller or snapshot: %v", err)
		}
		// Check Parse's copy separately without changing the canonical caller.
		separate := []byte(immutableLeafSource)
		other, err := losslessxml.Parse(separate)
		if err != nil {
			return err
		}
		separate[0] = '!'
		retained, err := other.Edit(nil, nil)
		if err != nil || !bytes.Equal(retained, []byte(immutableLeafSource)) {
			return fmt.Errorf("parsed snapshot aliases its caller: %v", err)
		}
		return nil
	})
}
