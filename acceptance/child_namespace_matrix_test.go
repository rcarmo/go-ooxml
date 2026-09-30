package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const childNamespaceMatrixCaseID = "@id-xml-go-child-namespace-matrix"

var childNamespaceSources = []string{
	`"<r/>"`,
	`"<r xmlns=\"u\"/>"`,
	`"<p:r xmlns:p=\"u\"/>"`,
	`"<r xmlns=\"u\" xmlns:n1=\"v\" xmlns:n2=\"occupied\"/>"`,
}

var childNamespaceURIs = []string{"", "u", "v", "fresh", "http://www.w3.org/XML/1998/namespace"}

func guardChildNamespaceMatrixCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"these four XML root sources, each interpreted as a JSON string",
		"these five namespace URI choices independently for each child element and flag attribute",
		`the XML editor inserts child with a flag attribute of JSON value "\t\r\n & 😀" and a plain grandchild of JSON text "x\ry\nz" for all 4 by 5 by 5 choices`,
		"reparsing preserves the root child and grandchild expanded names and the child attribute name and value for every choice",
		`the grandchild text equals JSON "x\ry\nz" for every choice`,
		"a separate empty edit returns each exact original root source",
	}
	if id != childNamespaceMatrixCaseID || line != lexicalEditingLine(52, 61) || p.Name != "Inserted element and attribute meanings survive a bounded namespace matrix" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected child-namespace matrix case %s %q at %d", id, p.Name, line)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("child-namespace matrix step %d drift", i+1)
		}
	}
	for i, want := range [][]string{append([]string{"source_json"}, childNamespaceSources...), append([]string{"namespace_uri"}, childNamespaceURIs...)} {
		arg := p.Steps[i].Argument
		if arg == nil || arg.DataTable == nil || len(arg.DataTable.Rows) != len(want) {
			return fmt.Errorf("child-namespace matrix table %d missing or changed", i+1)
		}
		for j, row := range arg.DataTable.Rows {
			if len(row.Cells) != 1 || row.Cells[0].Value != want[j] {
				return fmt.Errorf("child-namespace matrix table %d row %d drift", i+1, j)
			}
		}
	}
	for i := 2; i < len(p.Steps); i++ {
		if p.Steps[i].Argument != nil {
			return fmt.Errorf("child-namespace matrix unexpected step %d argument", i+1)
		}
	}
	if (len(childNamespaceSources) * len(childNamespaceURIs) * len(childNamespaceURIs)) != 100 {
		return fmt.Errorf("child-namespace matrix cardinality drift")
	}
	return nil
}

func TestChildNamespaceMatrixGuardRejectsDrift(t *testing.T) {
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
			if tag.Name == childNamespaceMatrixCaseID {
				if selected != nil {
					t.Fatal("duplicate matrix case")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardChildNamespaceMatrixCase(childNamespaceMatrixCaseID, selected, lexicalEditingLine(52, 61)) != nil {
		t.Fatal("canonical matrix guard failed")
	}
	copyCase := func() *messages.Pickle {
		data, err := json.Marshal(selected)
		if err != nil {
			t.Fatal(err)
		}
		var p messages.Pickle
		if err := json.Unmarshal(data, &p); err != nil {
			t.Fatal(err)
		}
		return &p
	}
	for _, tc := range []struct {
		name   string
		mutate func(*messages.Pickle)
	}{
		{"source row", func(p *messages.Pickle) { p.Steps[0].Argument.DataTable.Rows[2].Cells[0].Value = `"<r/>"` }},
		{"namespace row", func(p *messages.Pickle) { p.Steps[1].Argument.DataTable.Rows[4].Cells[0].Value = "u" }},
		{"missing row", func(p *messages.Pickle) { p.Steps[0].Argument.DataTable.Rows = p.Steps[0].Argument.DataTable.Rows[:4] }},
		{"step", func(p *messages.Pickle) { p.Steps[4].Text += " changed" }},
		{"example expansion", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "unexpected") }},
		{"extra argument", func(p *messages.Pickle) { p.Steps[3].Argument = p.Steps[0].Argument }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := copyCase()
			tc.mutate(p)
			if err := guardChildNamespaceMatrixCase(childNamespaceMatrixCaseID, p, lexicalEditingLine(52, 61)); err == nil {
				t.Fatal("guard accepted canonical drift")
			}
		})
	}
	if guardChildNamespaceMatrixCase("@id-xml-go-child-insertion-custody", selected, lexicalEditingLine(52, 61)) == nil || guardChildNamespaceMatrixCase(childNamespaceMatrixCaseID, selected, lexicalEditingLine(52, 61)+1) == nil {
		t.Fatal("guard accepted ID or line drift")
	}
}

func matrixColumn(table *godog.Table, header string, expected []string) ([]string, error) {
	if table == nil || len(table.Rows) != len(expected)+1 {
		return nil, fmt.Errorf("%s table row count drift", header)
	}
	values := make([]string, 0, len(expected))
	for i, row := range table.Rows {
		want := header
		if i > 0 {
			want = expected[i-1]
		}
		if len(row.Cells) != 1 || row.Cells[0].Value != want {
			return nil, fmt.Errorf("%s table row %d drift", header, i)
		}
		if i > 0 {
			values = append(values, row.Cells[0].Value)
		}
	}
	return values, nil
}

type matrixInsertion struct {
	doc      *losslessxml.Document
	caller   []byte
	original []byte
	output   []byte
	root     xml.Name
	child    xml.Name
	flag     xml.Name
}

func childNamespaceMatrixSteps(sc *godog.ScenarioContext) {
	var sources, uris []string
	var cases []matrixInsertion
	var attributeValue, grandchildText string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		sources, uris, cases = nil, nil, nil
		attributeValue, grandchildText = "", ""
		return ctx, nil
	})
	sc.Step(`^these four XML root sources, each interpreted as a JSON string$`, func(table *godog.Table) error {
		rows, err := matrixColumn(table, "source_json", childNamespaceSources)
		if err != nil {
			return err
		}
		for _, raw := range rows {
			var source string
			if err := json.Unmarshal([]byte(raw), &source); err != nil {
				return err
			}
			sources = append(sources, source)
		}
		return nil
	})
	sc.Step(`^these five namespace URI choices independently for each child element and flag attribute$`, func(table *godog.Table) error {
		var err error
		uris, err = matrixColumn(table, "namespace_uri", childNamespaceURIs)
		return err
	})
	sc.Step(`^the XML editor inserts child with a flag attribute of JSON value (".*") and a plain grandchild of JSON text (".*") for all 4 by 5 by 5 choices$`, func(attributeJSON, textJSON string) error {
		if len(sources) != 4 || len(uris) != 5 {
			return fmt.Errorf("child-namespace matrix missing source or namespace rows")
		}
		if err := json.Unmarshal([]byte(attributeJSON), &attributeValue); err != nil {
			return err
		}
		if err := json.Unmarshal([]byte(textJSON), &grandchildText); err != nil {
			return err
		}
		if attributeValue != "\t\r\n & 😀" || grandchildText != "x\ry\nz" {
			return fmt.Errorf("child-namespace matrix value drift")
		}
		for _, source := range sources {
			for _, elementURI := range uris {
				for _, attributeURI := range uris {
					caller := []byte(source)
					original := bytes.Clone(caller)
					doc, err := losslessxml.Parse(caller)
					if err != nil {
						return err
					}
					roots := doc.Elements()
					if len(roots) != 1 {
						return fmt.Errorf("matrix root count drift for %q", source)
					}
					child := xml.Name{Space: elementURI, Local: "child"}
					flag := xml.Name{Space: attributeURI, Local: "flag"}
					output, err := doc.InsertChildren([]losslessxml.ChildInsertion{{Parent: roots[0], Children: []losslessxml.NewElement{{Name: child, Attributes: []xml.Attr{{Name: flag, Value: attributeValue}}, Children: []losslessxml.NewElement{{Name: xml.Name{Local: "plain"}, Text: grandchildText}}}}}})
					if err != nil {
						return fmt.Errorf("source %q child %q attribute %q: %w", source, elementURI, attributeURI, err)
					}
					if !bytes.Equal(caller, original) {
						return fmt.Errorf("caller bytes changed for %q", source)
					}
					cases = append(cases, matrixInsertion{doc, caller, original, output, roots[0].Name(), child, flag})
				}
			}
		}
		if len(cases) != 100 {
			return fmt.Errorf("matrix executed %d of 100 choices", len(cases))
		}
		return nil
	})
	sc.Step(`^reparsing preserves the root child and grandchild expanded names and the child attribute name and value for every choice$`, func() error {
		if len(cases) != 100 {
			return fmt.Errorf("matrix only has %d choices", len(cases))
		}
		for i, choice := range cases {
			parsed, err := losslessxml.Parse(choice.output)
			if err != nil {
				return fmt.Errorf("choice %d: %w", i, err)
			}
			elements := parsed.Elements()
			if len(elements) != 3 || elements[0].Name() != choice.root || elements[1].Name() != choice.child || elements[2].Name() != (xml.Name{Local: "plain"}) {
				return fmt.Errorf("choice %d expanded element name/count drift", i)
			}
			for childIndex := 1; childIndex < 3; childIndex++ {
				parent, ok := elements[childIndex].Parent()
				if !ok || parent != elements[childIndex-1] {
					return fmt.Errorf("choice %d child %d parent drift", i, childIndex)
				}
			}
			attrs := elements[1].Attributes()
			if len(attrs) != 1 || attrs[0].Name != choice.flag || attrs[0].Value != attributeValue {
				return fmt.Errorf("choice %d expanded attribute/value drift", i)
			}
			if !bytes.Equal(choice.caller, choice.original) {
				return fmt.Errorf("choice %d caller bytes changed", i)
			}
		}
		return nil
	})
	sc.Step(`^the grandchild text equals JSON (".*") for every choice$`, func(textJSON string) error {
		var expected string
		if err := json.Unmarshal([]byte(textJSON), &expected); err != nil {
			return err
		}
		if expected != grandchildText || len(cases) != 100 {
			return fmt.Errorf("matrix grandchild text/cardinality drift")
		}
		for i, choice := range cases {
			parsed, err := losslessxml.Parse(choice.output)
			if err != nil {
				return err
			}
			text, leaf := parsed.Elements()[2].Text()
			if !leaf || text != expected {
				return fmt.Errorf("choice %d grandchild leaf/text drift: %q", i, text)
			}
		}
		return nil
	})
	sc.Step(`^a separate empty edit returns each exact original root source$`, func() error {
		if len(cases) != 100 {
			return fmt.Errorf("matrix empty edit only has %d choices", len(cases))
		}
		for i, choice := range cases {
			unchanged, err := choice.doc.Edit(nil, nil)
			if err != nil || !bytes.Equal(unchanged, choice.original) || !bytes.Equal(choice.caller, choice.original) {
				return fmt.Errorf("choice %d empty edit/caller changed: %v", i, err)
			}
		}
		return nil
	})
}
