package acceptance

import (
	"bytes"
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

const xmlExpandedAttributeCaseID = "@id-xml-expanded-attribute-lookup"
const xmlExpandedAttributeSourceJSON = `"<r xmlns=\"urn:default\" xmlns:a=\"urn:a\" xmlns:r=\"urn:a\" id=\"plain\" a:id=\"outer\"><child xmlns:r=\"urn:b\" r:id=\"inner\" xml:lang=\"en\"/><other r:id=\"sibling\"/></r>"`
const xmlExpandedAttributeInput = `<r xmlns="urn:default" xmlns:a="urn:a" xmlns:r="urn:a" id="plain" a:id="outer"><child xmlns:r="urn:b" r:id="inner" xml:lang="en"/><other r:id="sibling"/></r>`

// A seven-row canonical contract, including the absent and unqualified results.
var xmlExpandedAttributeRows = [][4]string{
	{"root", "id", "", `"plain"`},
	{"root", "id", "urn:default", "null"},
	{"root", "id", "urn:a", `"outer"`},
	{"child", "id", "urn:b", `"inner"`},
	{"child", "id", "urn:a", "null"},
	{"other", "id", "urn:a", `"sibling"`},
	{"child", "lang", "http://www.w3.org/XML/1998/namespace", `"en"`},
}

func xmlExpandedAttributeLine() int {
	if xmlLexicalCandidate() || batch2Candidate() {
		return 85
	}
	return 58
}

func guardXMLExpandedAttributeCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"XML values input encoded as JSON " + xmlExpandedAttributeSourceJSON,
		"the XML values input is parsed",
		"expanded attribute lookups return these JSON values",
	}
	if id != xmlExpandedAttributeCaseID || line != xmlExpandedAttributeLine() || p.Name != "Attribute lookup respects local prefix rebinding and unqualified names" || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("unexpected expanded-attribute case %s %q at %d", id, p.Name, line)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want {
			return fmt.Errorf("expanded-attribute step %d drift", i+1)
		}
		if i < 2 && p.Steps[i].Argument != nil {
			return fmt.Errorf("expanded-attribute step %d gained argument", i+1)
		}
	}
	return guardXMLExpandedAttributeTable(p.Steps[2].Argument)
}

func guardXMLExpandedAttributeTable(arg *messages.PickleStepArgument) error {
	if arg == nil || arg.DataTable == nil || len(arg.DataTable.Rows) != 8 {
		return fmt.Errorf("expanded-attribute table cardinality drift")
	}
	header := [4]string{"element", "local", "namespace", "value_json"}
	for i, row := range arg.DataTable.Rows {
		if len(row.Cells) != 4 {
			return fmt.Errorf("expanded-attribute table width drift at row %d", i)
		}
		want := header
		if i > 0 {
			want = xmlExpandedAttributeRows[i-1]
		}
		for col, cell := range row.Cells {
			if cell.Value != want[col] {
				return fmt.Errorf("expanded-attribute table row %d col %d drift", i, col)
			}
		}
	}
	return nil
}

func guardXMLExpandedAttributeRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "XML parsing and value inspection" {
		return fmt.Errorf("expanded-attribute feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "XML values, namespace lookup and safe escaping" {
			continue
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("expanded-attribute rule gained background")
			}
			if member.Scenario == nil {
				continue
			}
			for _, tag := range member.Scenario.Tags {
				if tag.Name == xmlExpandedAttributeCaseID {
					if len(member.Scenario.Tags) != 1 || int(tag.Location.Line) != xmlExpandedAttributeLine()-1 || len(member.Scenario.Examples) != 0 || len(member.Scenario.Steps) != 3 || int(member.Scenario.Location.Line) != xmlExpandedAttributeLine() {
						return fmt.Errorf("expanded-attribute canonical structure drift")
					}
					found++
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("expanded-attribute rule owns %d canonical scenarios", found)
	}
	return nil
}

func TestXMLExpandedAttributeGuardRejectsDrift(t *testing.T) {
	path := xmlParsingFeaturePath()
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
	if err := guardXMLExpandedAttributeRule(doc); err != nil {
		t.Fatal(err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == xmlExpandedAttributeCaseID {
				if selected != nil {
					t.Fatal("duplicate expanded-attribute pickle")
				}
				selected = p
			}
		}
	}
	if selected == nil || guardXMLExpandedAttributeCase(xmlExpandedAttributeCaseID, selected, xmlExpandedAttributeLine()) != nil {
		t.Fatal("canonical expanded-attribute case guard failed")
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
		{"source", func(p *messages.Pickle) { p.Steps[0].Text += " changed" }},
		{"parse", func(p *messages.Pickle) { p.Steps[1].Text += " changed" }},
		{"lookup", func(p *messages.Pickle) { p.Steps[2].Text += " changed" }},
		{"name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "unexpected") }},
		{"unexpected argument", func(p *messages.Pickle) { p.Steps[0].Argument = p.Steps[2].Argument }},
		{"missing argument", func(p *messages.Pickle) { p.Steps[2].Argument = nil }},
		{"row", func(p *messages.Pickle) { p.Steps[2].Argument.DataTable.Rows[4].Cells[2].Value = "urn:a" }},
		{"null type", func(p *messages.Pickle) { p.Steps[2].Argument.DataTable.Rows[2].Cells[3].Value = `""` }},
		{"missing row", func(p *messages.Pickle) { p.Steps[2].Argument.DataTable.Rows = p.Steps[2].Argument.DataTable.Rows[:7] }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := copyCase()
			tc.mutate(p)
			if guardXMLExpandedAttributeCase(xmlExpandedAttributeCaseID, p, xmlExpandedAttributeLine()) == nil {
				t.Fatal("guard accepted expanded-attribute drift")
			}
		})
	}
	if guardXMLExpandedAttributeCase(xmlEntityValuesCaseID, selected, xmlExpandedAttributeLine()) == nil || guardXMLExpandedAttributeCase(xmlExpandedAttributeCaseID, selected, xmlExpandedAttributeLine()-1) == nil {
		t.Fatal("guard accepted expanded-attribute ID or line drift")
	}
}

func expandedAttributeElements(doc *losslessxml.Document) (map[string]losslessxml.Element, error) {
	if doc == nil {
		return nil, fmt.Errorf("missing expanded-attribute document")
	}
	elements := doc.Elements()
	if len(elements) != 3 {
		return nil, fmt.Errorf("expanded-attribute element count drift: %d", len(elements))
	}
	names := []xml.Name{{Space: "urn:default", Local: "r"}, {Space: "urn:default", Local: "child"}, {Space: "urn:default", Local: "other"}}
	for i, e := range elements {
		if e.Name() != names[i] {
			return nil, fmt.Errorf("expanded-attribute element %d name drift: %v", i, e.Name())
		}
		parent, ok := e.Parent()
		if i == 0 && ok || i > 0 && (!ok || parent.Ordinal() != elements[0].Ordinal() || !bytes.Equal(parent.Raw(), elements[0].Raw())) {
			return nil, fmt.Errorf("expanded-attribute element %d parent drift", i)
		}
	}
	rootNS, childNS, otherNS := elements[0].Namespaces(), elements[1].Namespaces(), elements[2].Namespaces()
	if rootNS["r"] != "urn:a" || childNS["r"] != "urn:b" || otherNS["r"] != "urn:a" {
		return nil, fmt.Errorf("expanded-attribute local/sibling namespace scope drift")
	}
	return map[string]losslessxml.Element{"root": elements[0], "child": elements[1], "other": elements[2]}, nil
}

func assertExpandedAttributeRows(doc *losslessxml.Document, table *godog.Table) error {
	if table == nil || len(table.Rows) != 8 {
		return fmt.Errorf("expanded-attribute runtime table size drift")
	}
	elements, err := expandedAttributeElements(doc)
	if err != nil {
		return err
	}
	seen := map[[3]string]bool{}
	for i, row := range table.Rows {
		if len(row.Cells) != 4 {
			return fmt.Errorf("expanded-attribute runtime row width drift")
		}
		if i == 0 {
			for j, col := range [4]string{"element", "local", "namespace", "value_json"} {
				if row.Cells[j].Value != col {
					return fmt.Errorf("expanded-attribute runtime header drift")
				}
			}
			continue
		}
		element, local, ns, raw := row.Cells[0].Value, row.Cells[1].Value, row.Cells[2].Value, row.Cells[3].Value
		if [4]string{element, local, ns, raw} != xmlExpandedAttributeRows[i-1] {
			return fmt.Errorf("expanded-attribute runtime row %d drift", i)
		}
		key := [3]string{element, local, ns}
		if seen[key] {
			return fmt.Errorf("expanded-attribute duplicate lookup key %v", key)
		}
		seen[key] = true
		target, ok := elements[element]
		if !ok {
			return fmt.Errorf("expanded-attribute unknown element %s", element)
		}
		matches := 0
		var got string
		for _, attr := range target.Attributes() {
			if attr.Name == (xml.Name{Space: ns, Local: local}) {
				matches++
				got = attr.Value
			}
		}
		if matches > 1 {
			return fmt.Errorf("expanded-attribute ambiguous match for %v", key)
		}
		if raw == "null" {
			if matches != 0 {
				return fmt.Errorf("expanded-attribute absent lookup %v returned %q", key, got)
			}
			continue
		}
		var expected string
		if err := json.Unmarshal([]byte(raw), &expected); err != nil || matches != 1 || got != expected {
			return fmt.Errorf("expanded-attribute lookup %v = %q (%d), want %s: %v", key, got, matches, raw, err)
		}
	}
	if len(seen) != len(xmlExpandedAttributeRows) {
		return fmt.Errorf("expanded-attribute row coverage drift")
	}
	return nil
}

func registerXMLExpandedAttributeSteps(sc *godog.ScenarioContext, caller, original *[]byte, doc **losslessxml.Document) {
	sc.Step(`^expanded attribute lookups return these JSON values$`, func(table *godog.Table) error {
		if !bytes.Equal(*original, []byte(xmlExpandedAttributeInput)) || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("expanded-attribute source or caller drift")
		}
		if err := assertExpandedAttributeRows(*doc, table); err != nil {
			return err
		}
		root := (*doc).Elements()[0]
		attrs := root.Attributes()
		if len(attrs) != 2 || !bytes.Equal(root.Raw(), []byte(xmlExpandedAttributeInput)) {
			return fmt.Errorf("expanded-attribute root bytes/attribute count drift")
		}
		attrs[0].Value = "corrupt"
		if root.Attributes()[0].Value != "plain" {
			return fmt.Errorf("expanded-attribute copied attributes alias snapshot")
		}
		out, err := (*doc).Edit(nil, nil)
		if err != nil || !bytes.Equal(out, *original) || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("expanded-attribute empty edit/source drift: %v", err)
		}
		out[0] = '!'
		again, err := (*doc).Edit(nil, nil)
		if err != nil || !bytes.Equal(again, *original) || !bytes.Equal(*caller, *original) {
			return fmt.Errorf("expanded-attribute no-op output aliases source/caller: %v", err)
		}
		(*caller)[0] = '!'
		third, err := (*doc).Edit(nil, nil)
		if err != nil || !bytes.Equal(third, *original) {
			return fmt.Errorf("expanded-attribute snapshot aliases caller: %v", err)
		}
		return nil
	})
}

func TestXMLExpandedAttributeRefusalAndRetry(t *testing.T) {
	bad := []byte(`<r xmlns:a="urn:a" xmlns:b="urn:a" a:id="one" b:id="two"/>`)
	before := bytes.Clone(bad)
	invalid, err := losslessxml.Parse(bad)
	if err == nil || invalid != nil || !bytes.Equal(bad, before) {
		t.Fatalf("duplicate expanded attribute accepted or changed caller: %v", err)
	}
	good := []byte(xmlExpandedAttributeInput)
	original := bytes.Clone(good)
	doc, err := losslessxml.Parse(good)
	if err != nil || !bytes.Equal(good, original) {
		t.Fatalf("expanded-attribute valid retry failed or changed caller: %v", err)
	}
	elements, err := expandedAttributeElements(doc)
	if err != nil {
		t.Fatal(err)
	}
	if attrs := elements["child"].Attributes(); len(attrs) != 2 || attrs[0].Name != (xml.Name{Space: "urn:b", Local: "id"}) || attrs[0].Value != "inner" {
		t.Fatal("expanded-attribute valid retry lost child binding")
	}
}
