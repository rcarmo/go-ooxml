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

const attributeSpliceSource = `<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="old"/><t>text</t></r>`

var attributeSpliceRows = map[string]struct {
	value, output string
	name          xml.Name
	attributes    [3]string // decoded a, p:n, fresh (empty if absent)
}{
	"a":     {"x'y&z", `<r xmlns:p="urn:p"><t a = 'x&#39;y&amp;z' p:n="old"/><t>text</t></r>`, xml.Name{Local: "a"}, [3]string{"x'y&z", "old", ""}},
	"p:n":   {"new", `<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="new"/><t>text</t></r>`, xml.Name{Space: "urn:p", Local: "n"}, [3]string{"a&b", "new", ""}},
	"fresh": {"value", `<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="old" fresh="value"/><t>text</t></r>`, xml.Name{Local: "fresh"}, [3]string{"a&b", "old", "value"}},
}

// The feature is a Scenario Outline; bind only these exact three expanded rows.
func guardAttributeSpliceCase(id string, p *messages.Pickle, line int) (string, error) {
	const scenario = "Update one attribute without reserialising its neighbours"
	if id != attributeSpliceCaseID || p.Name != scenario || len(p.Steps) != 3 || len(p.AstNodeIds) != 2 {
		return "", fmt.Errorf("unexpected attribute-splice case %s %q", id, p.Name)
	}
	for name, row := range attributeSpliceRows {
		if p.Steps[1].Text != "a lexical edit sets the attribute "+name+" of the first t element to "+row.value {
			continue
		}
		base := 27
		if xmlLexicalCandidate() {
			base = 36
		}
		rowLine := map[string]int{"a": base, "p:n": base + 1, "fresh": base + 2}[name]
		if line != rowLine {
			return "", fmt.Errorf("attribute-splice row %s moved to %d", name, line)
		}
		want := []string{
			"the XML source is " + attributeSpliceSource,
			"a lexical edit sets the attribute " + name + " of the first t element to " + row.value,
			"the complete output bytes equal " + row.output,
		}
		for i, text := range want {
			if p.Steps[i].Text != text {
				return "", fmt.Errorf("attribute-splice row %s step %d drift", name, i+1)
			}
		}
		return name, nil
	}
	return "", fmt.Errorf("unexpected attribute-splice row %q", p.Name)
}

func attributeSpliceSteps(sc *godog.ScenarioContext) {
	var caller, original, output []byte
	var doc *losslessxml.Document
	var selected string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		caller, original, output, doc, selected = nil, nil, nil, nil, ""
		return ctx, nil
	})
	sc.Step(`^the XML source is (<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="old"/><t>text</t></r>)$`, func(source string) error {
		if source != attributeSpliceSource {
			return fmt.Errorf("source drift")
		}
		caller = []byte(source)
		original = bytes.Clone(caller)
		return nil
	})
	sc.Step(`^a lexical edit sets the attribute (a|p:n|fresh) of the first t element to (.*)$`, func(name, value string) error {
		row, ok := attributeSpliceRows[name]
		if !ok || value != row.value || !bytes.Equal(caller, []byte(attributeSpliceSource)) {
			return fmt.Errorf("unexpected attribute row %s=%q", name, value)
		}
		var err error
		doc, err = losslessxml.Parse(caller)
		if err != nil {
			return err
		}
		if err = checkAttributeSpliceDocument(doc, [3]string{"a&b", "old", ""}); err != nil {
			return fmt.Errorf("input: %w", err)
		}
		output, err = doc.Edit(nil, []losslessxml.AttributeEdit{{Target: doc.Elements()[1], Name: row.name, Value: value}})
		selected = name
		return err
	})
	sc.Step(`^the complete output bytes equal (<r xmlns:p="urn:p"><t .*)$`, func(expected string) error {
		row, ok := attributeSpliceRows[selected]
		if !ok || expected != row.output || !bytes.Equal(output, []byte(row.output)) || !bytes.Equal(caller, original) || !bytes.Equal(caller, []byte(attributeSpliceSource)) {
			return fmt.Errorf("attribute-splice %s bytes or caller changed: %q", selected, output)
		}
		edited, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		if err := checkAttributeSpliceDocument(edited, row.attributes); err != nil {
			return fmt.Errorf("edited: %w", err)
		}
		empty, err := doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(empty, []byte(attributeSpliceSource)) {
			return fmt.Errorf("empty edit changed source: %v", err)
		}
		same, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "a&b"}})
		if err != nil || !bytes.Equal(same, []byte(attributeSpliceSource)) {
			return fmt.Errorf("same-value edit changed original lexical bytes: %v", err)
		}
		refused, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "x"}, {Target: doc.Elements()[1], Name: xml.Name{Local: "a"}, Value: "y"}})
		if err == nil || refused != nil {
			return fmt.Errorf("duplicate batch accepted or emitted bytes")
		}
		again, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: doc.Elements()[1], Name: row.name, Value: row.value}})
		if err != nil || !bytes.Equal(again, []byte(row.output)) {
			return fmt.Errorf("refusal consumed snapshot: %v", err)
		}
		output[0] = '!'
		retained, err := doc.Edit(nil, nil)
		if err != nil || !bytes.Equal(retained, []byte(attributeSpliceSource)) || !bytes.Equal(caller, original) {
			return fmt.Errorf("returned bytes alias caller or snapshot: %v", err)
		}
		return nil
	})
}

func checkAttributeSpliceDocument(doc *losslessxml.Document, want [3]string) error {
	elements := doc.Elements()
	if len(elements) != 3 || elements[0].Name() != (xml.Name{Local: "r"}) || elements[1].Name() != (xml.Name{Local: "t"}) || elements[2].Name() != (xml.Name{Local: "t"}) {
		return fmt.Errorf("wrong root or t elements")
	}
	if text, leaf := elements[2].Text(); !leaf || text != "text" || !bytes.Equal(elements[2].Raw(), []byte(`<t>text</t>`)) {
		return fmt.Errorf("untouched sibling changed")
	}
	got := [3]string{}
	present := [3]bool{}
	for _, a := range elements[1].Attributes() {
		switch a.Name {
		case (xml.Name{Local: "a"}):
			got[0], present[0] = a.Value, true
		case (xml.Name{Space: "urn:p", Local: "n"}):
			got[1], present[1] = a.Value, true
		case (xml.Name{Local: "fresh"}):
			got[2], present[2] = a.Value, true
		default:
			return fmt.Errorf("unexpected attribute %v", a.Name)
		}
	}
	if got != want || !present[0] || !present[1] || present[2] != (want[2] != "") {
		return fmt.Errorf("attribute values %+v, present %+v", got, present)
	}
	return nil
}
