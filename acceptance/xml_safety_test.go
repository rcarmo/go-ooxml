package acceptance

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/xmlsnapshot"
)

const (
	xmlPrototypeCaseID     = "@id-xml-prototype-safe-attributes"
	xmlNamespaceCaseID     = "@id-xml-immutable-namespace-metadata"
	xmlMalformedCaseID     = "@id-xml-typed-parse-error"
	xmlPrototypeSourceJSON = `"<r __proto__=\"polluted\" constructor=\"safe\"/>"`
	xmlNamespaceSourceJSON = `"<r xmlns:a=\"urn:a\" a:id=\"outer\"/>"`
	xmlMalformedSourceJSON = `"<a></b>"`
	xmlPrototypeInput      = `<r __proto__="polluted" constructor="safe"/>`
	xmlNamespaceInput      = `<r xmlns:a="urn:a" a:id="outer"/>`
	xmlMalformedInput      = `<a></b>`
)

type xmlSafetyState struct {
	caller, original []byte
	doc              *xmlsnapshot.Document
	err              error
	metadata         map[string]string
}

func xmlSafetyLine(id string) int {
	if xmlLexicalCandidate() || postBatch2Reference() {
		return map[string]int{xmlPrototypeCaseID: 107, xmlNamespaceCaseID: 115, xmlMalformedCaseID: 146}[id]
	}
	return map[string]int{xmlPrototypeCaseID: 80, xmlNamespaceCaseID: 88, xmlMalformedCaseID: 119}[id]
}

func guardXMLSafetyCase(id string, p *messages.Pickle, line int) error {
	var name string
	var steps []string
	switch id {
	case xmlPrototypeCaseID:
		name = "Special-looking attribute names remain ordinary XML data"
		steps = []string{
			"XML values input encoded as JSON " + xmlPrototypeSourceJSON,
			"the XML values input is parsed",
			"the root attributes are exactly __proto__=polluted and constructor=safe",
			"the root remains r with no text and no children",
			"fresh attribute reads and the original XML source are unchanged",
		}
	case xmlNamespaceCaseID:
		name = "Returned namespace metadata cannot corrupt the parsed XML model"
		steps = []string{
			"XML values input encoded as JSON " + xmlNamespaceSourceJSON,
			"the XML values input is parsed",
			"the namespace recorded for a:id is urn:a",
			"changing a returned namespace snapshot to urn:changed and deleting a:id are attempted",
			"fresh namespace and attribute reads still return urn:a and outer",
			"the parsed root structure and original XML source are unchanged",
		}
	case xmlMalformedCaseID:
		name = "Malformed XML returns a documented failure category without a partial document"
		steps = []string{
			"XML values input encoded as JSON " + xmlMalformedSourceJSON,
			"the XML values input is parsed",
			"parsing fails with category malformed-xml and no document result",
			"the original XML source is unchanged",
		}
	default:
		return fmt.Errorf("unexpected XML safety ID %s", id)
	}
	if line != xmlSafetyLine(id) || p.Name != name || len(p.AstNodeIds) != 1 || len(p.Steps) != len(steps) {
		return fmt.Errorf("XML safety scenario %s %q at line %d drift", id, p.Name, line)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("XML safety step %s #%d drift", id, i+1)
		}
	}
	return nil
}

func guardXMLSafetyRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "XML parsing and value inspection" {
		return fmt.Errorf("XML safety feature drift")
	}
	seen := map[string]bool{}
	for _, child := range doc.Feature.Children {
		if child.Rule == nil || child.Rule.Name != "XML values, namespace lookup and safe escaping" {
			continue
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				return fmt.Errorf("XML safety rule gained background")
			}
			if member.Scenario == nil {
				continue
			}
			s := member.Scenario
			for _, tag := range s.Tags {
				var id string
				var expectedTagLine, expectedLine, expectedSteps int
				switch tag.Name {
				case xmlPrototypeCaseID:
					id, expectedLine, expectedSteps = tag.Name, xmlSafetyLine(tag.Name), 5
				case xmlNamespaceCaseID:
					id, expectedLine, expectedSteps = tag.Name, xmlSafetyLine(tag.Name), 6
				case xmlMalformedCaseID:
					id, expectedLine, expectedSteps = tag.Name, xmlSafetyLine(tag.Name), 4
				default:
					continue
				}
				expectedTagLine = expectedLine - 1
				if seen[id] || len(s.Tags) != 2 || s.Tags[0].Name != map[string]string{xmlMalformedCaseID: "@profile-xml-failure-category", xmlPrototypeCaseID: "@profile-xml-model-safety", xmlNamespaceCaseID: "@profile-xml-model-safety"}[id] || int(tag.Location.Line) != expectedTagLine || int(s.Location.Line) != expectedLine || len(s.Examples) != 0 || len(s.Steps) != expectedSteps {
					return fmt.Errorf("XML safety structure drift: %s tags=%s,%s tagLine=%d scenarioLine=%d examples=%d steps=%d", id, s.Tags[0].Name, s.Tags[1].Name, tag.Location.Line, s.Location.Line, len(s.Examples), len(s.Steps))
				}
				seen[id] = true
			}
		}
	}
	if len(seen) != 3 {
		return fmt.Errorf("XML safety rule owns %d of 3 cases", len(seen))
	}
	return nil
}

// Published v0.152.0 retains the old JavaScript-specific profile. The
// generalized shared feature activates these native Go bindings. The runner
// separately verifies the release or explicit candidate pin and checkout.
func xmlSafetyCandidate() bool {
	data, err := os.ReadFile(xmlParsingFeaturePath())
	return err == nil && bytes.Contains(data, []byte("@profile-xml-model-safety @id-xml-prototype-safe-attributes")) &&
		bytes.Contains(data, []byte("@profile-xml-failure-category @id-xml-typed-parse-error"))
}

func TestXMLSafetyGuardRejectsDrift(t *testing.T) {
	loadReferencePin(t)
	if !xmlSafetyCandidate() {
		t.Skip("XML safety native binding requires explicit sealed candidate pin and root")
	}
	path := xmlParsingFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil || guardXMLSafetyRule(doc) != nil {
		t.Fatalf("canonical XML safety Rule drift: %v, %v", err, guardXMLSafetyRule(doc))
	}
	seen := map[string]bool{}
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			line := xmlSafetyLine(tag.Name)
			if line == 0 {
				continue
			}
			if seen[tag.Name] || guardXMLSafetyCase(tag.Name, p, line) != nil {
				t.Fatalf("canonical XML safety case guard failed: %s", tag.Name)
			}
			seen[tag.Name] = true
			clone := *p
			clone.Steps = append([]*messages.PickleStep(nil), p.Steps...)
			modified := *clone.Steps[2]
			modified.Text += " changed"
			clone.Steps[2] = &modified
			if guardXMLSafetyCase(tag.Name, &clone, line) == nil || guardXMLSafetyCase(tag.Name, p, line+1) == nil {
				t.Fatalf("XML safety guard accepted drift: %s", tag.Name)
			}
		}
	}
	if len(seen) != 3 {
		t.Fatalf("XML safety cases found: %d", len(seen))
	}
}

func xmlSafetySteps(sc *godog.ScenarioContext, s *xmlSafetyState) {
	root := func() (xmlsnapshot.Element, error) {
		if s.doc == nil || s.err != nil || !bytes.Equal(s.caller, s.original) {
			return xmlsnapshot.Element{}, fmt.Errorf("no intact parsed XML snapshot")
		}
		return s.doc.Root(), nil
	}
	structure := func() error {
		r, err := root()
		if err != nil {
			return err
		}
		text, _ := r.Text()
		if r.Name() != (xml.Name{Local: "r"}) || text != "" || len(r.Children()) != 0 || !bytes.Equal(s.doc.Source(), s.original) {
			return fmt.Errorf("XML safety root/source drift")
		}
		return nil
	}
	sc.Step(`^the root attributes are exactly __proto__=polluted and constructor=safe$`, func() error {
		r, err := root()
		if err != nil {
			return err
		}
		attrs := r.Attributes()
		if len(attrs) != 2 || attrs[0] != (xml.Attr{Name: xml.Name{Local: "__proto__"}, Value: "polluted"}) || attrs[1] != (xml.Attr{Name: xml.Name{Local: "constructor"}, Value: "safe"}) {
			return fmt.Errorf("XML literal attribute set drift: %v", attrs)
		}
		attrs[0].Value = "changed"
		return nil
	})
	sc.Step(`^the root remains r with no text and no children$`, structure)
	sc.Step(`^fresh attribute reads and the original XML source are unchanged$`, func() error {
		r, err := root()
		if err != nil {
			return err
		}
		attrs := r.Attributes()
		if len(attrs) != 2 || attrs[0].Value != "polluted" || attrs[1].Value != "safe" || !bytes.Equal(s.caller, s.original) {
			return fmt.Errorf("fresh XML attribute/source drift: %v", attrs)
		}
		return structure()
	})
	sc.Step(`^the namespace recorded for a:id is urn:a$`, func() error {
		r, err := root()
		if err != nil {
			return err
		}
		s.metadata = r.AttributeNamespaces()
		value, ok := r.Attribute("urn:a", "id")
		if len(s.metadata) != 1 || s.metadata["a:id"] != "urn:a" || !ok || value != "outer" {
			return fmt.Errorf("XML namespace/attribute drift: %v, %q", s.metadata, value)
		}
		return nil
	})
	sc.Step(`^changing a returned namespace snapshot to urn:changed and deleting a:id are attempted$`, func() error {
		if s.metadata == nil {
			return fmt.Errorf("no production metadata")
		}
		s.metadata["a:id"] = "urn:changed"
		r, err := root()
		if err != nil {
			return err
		}
		value, ok := r.Attribute("urn:a", "id")
		fresh := r.AttributeNamespaces()
		if fresh["a:id"] != "urn:a" || !ok || value != "outer" {
			return fmt.Errorf("namespace mutation reached original model: %v", fresh)
		}
		delete(s.metadata, "a:id")
		value, ok = r.Attribute("urn:a", "id")
		fresh = r.AttributeNamespaces()
		if fresh["a:id"] != "urn:a" || !ok || value != "outer" {
			return fmt.Errorf("namespace deletion reached original model: %v", fresh)
		}
		return nil
	})
	sc.Step(`^fresh namespace and attribute reads still return urn:a and outer$`, func() error {
		r, err := root()
		if err != nil {
			return err
		}
		metadata := r.AttributeNamespaces()
		value, ok := r.Attribute("urn:a", "id")
		if len(metadata) != 1 || metadata["a:id"] != "urn:a" || !ok || value != "outer" {
			return fmt.Errorf("fresh namespace/attribute drift: %v, %q", metadata, value)
		}
		return nil
	})
	sc.Step(`^the parsed root structure and original XML source are unchanged$`, func() error {
		if !bytes.Equal(s.caller, s.original) {
			return fmt.Errorf("XML caller source changed")
		}
		return structure()
	})
	sc.Step(`^parsing fails with category malformed-xml and no document result$`, func() error {
		var classified *xmlsnapshot.ParseError
		if s.doc != nil || !errors.As(s.err, &classified) || classified.Category != xmlsnapshot.CategoryMalformedXML {
			return fmt.Errorf("missing typed XML failure: doc=%v error=%v", s.doc, s.err)
		}
		// A valid source must parse; an unrelated failure must not share the category.
		good, goodErr := xmlsnapshot.Parse([]byte("<a/>"))
		other, otherErr := xmlsnapshot.Parse([]byte("<!DOCTYPE a><a/>"))
		var unrelated *xmlsnapshot.ParseError
		if good == nil || goodErr != nil || good.Root().Name() != (xml.Name{Local: "a"}) || other != nil || otherErr == nil || errors.As(otherErr, &unrelated) {
			return fmt.Errorf("XML failure positive/negative controls: good=%v/%v unrelated=%v/%v", good, goodErr, other, otherErr)
		}
		return nil
	})
	sc.Step(`^the original XML source is unchanged$`, func() error {
		if !bytes.Equal(s.caller, s.original) || string(s.original) != xmlMalformedInput {
			return fmt.Errorf("malformed XML caller source changed")
		}
		return nil
	})
}
