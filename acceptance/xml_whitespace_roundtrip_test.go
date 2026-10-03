package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"regexp"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/utils"
)

const xmlWhitespaceCaseID = "@id-xml-escaping-whitespace-roundtrip"
const xmlWhitespaceJSON = `"x\r\n\ty"`

func xmlWhitespaceLine() int {
	if xmlLexicalCandidate() || historicalAPI18Shape() {
		return 140
	}
	if xmlSafetyCandidate() {
		return 113
	}
	return 111
}

var xmlWhitespaceSteps = []string{
	"an XML escaping value encoded as JSON " + xmlWhitespaceJSON,
	"the value is escaped separately as text and as an attribute and both are parsed",
	"the decoded text and attribute both equal JSON " + xmlWhitespaceJSON,
}

// Reject drift in the exact three-step canonical selection, not its planned siblings.
func guardXMLWhitespaceRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "XML parsing and value inspection" || len(doc.Feature.Tags) != 1 || doc.Feature.Tags[0].Name != "@planned" {
		return fmt.Errorf("XML whitespace feature drift")
	}
	found := 0
	for _, child := range doc.Feature.Children {
		if child.Background != nil {
			return fmt.Errorf("XML whitespace feature background drift")
		}
		if child.Rule == nil {
			if child.Scenario != nil {
				for _, tag := range child.Scenario.Tags {
					if tag.Name == xmlWhitespaceCaseID {
						return fmt.Errorf("XML whitespace scenario escaped its rule")
					}
				}
			}
			continue
		}
		rule := child.Rule
		for _, member := range rule.Children {
			if member.Background != nil && rule.Name == "XML values, namespace lookup and safe escaping" {
				return fmt.Errorf("XML whitespace rule background drift")
			}
			if member.Scenario == nil {
				continue
			}
			for _, tag := range member.Scenario.Tags {
				if tag.Name != xmlWhitespaceCaseID {
					continue
				}
				found++
				s := member.Scenario
				if rule.Name != "XML values, namespace lookup and safe escaping" || len(rule.Tags) != 0 || len(s.Tags) != 1 || int(tag.Location.Line) != xmlWhitespaceLine()-1 || int(s.Location.Line) != xmlWhitespaceLine() || s.Keyword != "Scenario" || s.Name != "Escaped whitespace survives text and attribute parsing" || len(s.Examples) != 0 || len(s.Steps) != 3 {
					return fmt.Errorf("XML whitespace scenario structure drift")
				}
				for i, keyword := range []string{"Given ", "When ", "Then "} {
					if s.Steps[i].Keyword != keyword || int(s.Steps[i].Location.Line) != xmlWhitespaceLine()+1+i || s.Steps[i].Text != xmlWhitespaceSteps[i] || s.Steps[i].DocString != nil || s.Steps[i].DataTable != nil {
						return fmt.Errorf("XML whitespace authored step %d drift", i+1)
					}
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("XML whitespace scenario count %d", found)
	}
	return nil
}

func guardXMLWhitespaceCase(id string, p *messages.Pickle, line int) error {
	if id != xmlWhitespaceCaseID || p == nil || line != xmlWhitespaceLine() || p.Name != "Escaped whitespace survives text and attribute parsing" || len(p.AstNodeIds) != 1 || len(p.Tags) != 2 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != id || len(p.Steps) != len(xmlWhitespaceSteps) {
		return fmt.Errorf("XML whitespace pickle identity drift")
	}
	for i, step := range xmlWhitespaceSteps {
		if p.Steps[i].Text != step || p.Steps[i].Argument != nil {
			return fmt.Errorf("XML whitespace step %d drift", i+1)
		}
	}
	return nil
}

func TestXMLWhitespaceGuardRejectsDrift(t *testing.T) {
	path := xmlParsingFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil || guardXMLWhitespaceRule(doc) != nil {
		t.Fatalf("canonical XML whitespace rule drift: %v", err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == xmlWhitespaceCaseID {
				if selected != nil {
					t.Fatal("duplicate whitespace pickle")
				}
				selected = p
			}
		}
	}
	if err := guardXMLWhitespaceCase(xmlWhitespaceCaseID, selected, xmlWhitespaceLine()); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name   string
		change func(*messages.Pickle)
	}{
		{"name", func(p *messages.Pickle) { p.Name += " altered" }},
		{"tag", func(p *messages.Pickle) { p.Tags[1].Name = "@id-other" }},
		{"extra tag", func(p *messages.Pickle) { p.Tags = append(p.Tags, &messages.PickleTag{Name: "@extra"}) }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "row") }},
		{"argument", func(p *messages.Pickle) { p.Steps[1].Argument = &messages.PickleStepArgument{} }},
		{"given", func(p *messages.Pickle) { p.Steps[0].Text += " altered" }},
		{"when", func(p *messages.Pickle) { p.Steps[1].Text += " altered" }},
		{"then", func(p *messages.Pickle) { p.Steps[2].Text += " altered" }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := *selected
			p.AstNodeIds = append([]string(nil), selected.AstNodeIds...)
			p.Tags = append([]*messages.PickleTag(nil), selected.Tags...)
			for i, tag := range p.Tags {
				copy := *tag
				p.Tags[i] = &copy
			}
			p.Steps = append([]*messages.PickleStep(nil), selected.Steps...)
			for i, step := range p.Steps {
				copy := *step
				p.Steps[i] = &copy
			}
			tc.change(&p)
			if guardXMLWhitespaceCase(xmlWhitespaceCaseID, &p, xmlWhitespaceLine()) == nil {
				t.Fatal("accepted canonical drift")
			}
		})
	}
	if guardXMLWhitespaceCase(xmlWhitespaceCaseID, selected, xmlWhitespaceLine()+1) == nil {
		t.Fatal("accepted moved scenario")
	}
	for _, tc := range []struct {
		name   string
		change func(*messages.GherkinDocument)
	}{
		{"feature", func(d *messages.GherkinDocument) { d.Feature.Name += " altered" }},
		{"feature tag", func(d *messages.GherkinDocument) { d.Feature.Tags[0].Name = "@implemented" }},
		{"rule", func(d *messages.GherkinDocument) { d.Feature.Children[1].Rule.Name += " altered" }},
		{"scenario tag line", func(d *messages.GherkinDocument) { findWhitespaceScenario(d).Tags[0].Location.Line++ }},
		{"authored keyword", func(d *messages.GherkinDocument) { findWhitespaceScenario(d).Steps[1].Keyword = "And " }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f, err := os.Open(path)
			if err != nil {
				t.Fatal(err)
			}
			defer f.Close()
			other, err := gherkin.ParseGherkinDocument(f, next)
			if err != nil {
				t.Fatal(err)
			}
			tc.change(other)
			if guardXMLWhitespaceRule(other) == nil {
				t.Fatal("accepted source structure drift")
			}
		})
	}
}

func findWhitespaceScenario(doc *messages.GherkinDocument) *messages.Scenario {
	for _, c := range doc.Feature.Children {
		if c.Rule == nil {
			continue
		}
		for _, m := range c.Rule.Children {
			if m.Scenario != nil && len(m.Scenario.Tags) == 1 && m.Scenario.Tags[0].Name == xmlWhitespaceCaseID {
				return m.Scenario
			}
		}
	}
	return nil
}

type whitespaceValue struct {
	XMLName   xml.Name `xml:"r"`
	Attribute string   `xml:"a,attr"`
	Text      string   `xml:",chardata"`
}

var whitespaceMarkup = regexp.MustCompile(`(?s)^(?:<\?xml[^?]*\?>\s*)?<r a="([^"]*)">([^<]*)</r>$`)
var whitespaceCR = regexp.MustCompile(`&#(?:0*13|[xX]0*[dD]);`)
var whitespaceLF = regexp.MustCompile(`&#(?:0*10|[xX]0*[aA]);`)
var whitespaceTAB = regexp.MustCompile(`&#(?:0*9|[xX]0*9);`)

// A whole-element writer call gives the same caller-owned value to both XML
// contexts. Only the public Go serializer and reparser determine the result.
func writeWhitespace(value string) ([]byte, error) {
	return utils.MarshalXMLWithHeader(whitespaceValue{Attribute: value, Text: value})
}
func readWhitespace(b []byte) (whitespaceValue, error) {
	var out whitespaceValue
	err := utils.UnmarshalXML(b, &out)
	if err == nil && (out.XMLName != (xml.Name{Local: "r"}) || !bytes.Contains(b, []byte(` a="`))) {
		err = fmt.Errorf("unexpected XML structure")
	}
	return out, err
}

// These are semantic reference-presence checks, not generated-output equality.
// Replacing one reference with a DIFFERENT valid character permits the public
// reparser to observe which of the two contexts was damaged.
func whitespaceReferenceRegions(b []byte) (attr, text []byte, err error) {
	match := whitespaceMarkup.FindSubmatch(b)
	if len(match) != 3 {
		return nil, nil, fmt.Errorf("writer did not produce one r element with a and text")
	}
	attr, text = match[1], match[2]
	if !whitespaceCR.Match(attr) || !whitespaceLF.Match(attr) || !whitespaceTAB.Match(attr) || !whitespaceCR.Match(text) {
		return nil, nil, fmt.Errorf("writer omitted required whitespace references")
	}
	return attr, text, nil
}

func perturbWhitespaceContext(b []byte, context string, reference *regexp.Regexp) ([]byte, error) {
	indices := whitespaceMarkup.FindSubmatchIndex(b)
	if len(indices) != 6 {
		return nil, fmt.Errorf("missing XML regions")
	}
	start, end := indices[2], indices[3]
	if context == "text" {
		start, end = indices[4], indices[5]
	} else if context != "attribute" {
		return nil, fmt.Errorf("unknown context %s", context)
	}
	loc := reference.FindIndex(b[start:end])
	if loc == nil {
		return nil, fmt.Errorf("missing whitespace reference in %s", context)
	}
	left, right := start+loc[0], start+loc[1]
	out := append([]byte(nil), b[:left]...)
	out = append(out, []byte("&#90;")...) // 'Z' is valid XML but not the supplied CR.
	return append(out, b[right:]...), nil
}

func assertWhitespaceRoundtrip(encoded []byte, want string) error {
	if _, _, err := whitespaceReferenceRegions(encoded); err != nil {
		return err
	}
	got, err := readWhitespace(encoded)
	if err != nil {
		return err
	}
	if got.Attribute != want || got.Text != want {
		return fmt.Errorf("writer/reparser lost whitespace: text=%q attribute=%q want=%q", got.Text, got.Attribute, want)
	}
	return nil
}

func TestXMLWhitespacePublicControls(t *testing.T) {
	for _, value := range []string{"ordinary content", "x\r\n\ty"} {
		b, err := writeWhitespace(value)
		if err != nil {
			t.Fatal(err)
		}
		got, err := readWhitespace(b)
		if err != nil || got.Text != value || got.Attribute != value {
			t.Fatalf("public writer/reparser: text=%q attr=%q err=%v", got.Text, got.Attribute, err)
		}
		if value == "ordinary content" {
			continue
		}
		if err := assertWhitespaceRoundtrip(b, value); err != nil {
			t.Fatal(err)
		}
		for _, tc := range []struct {
			name, context string
			reference     *regexp.Regexp
		}{
			{"text-CR", "text", whitespaceCR},
			{"attribute-CR", "attribute", whitespaceCR},
			{"attribute-LF", "attribute", whitespaceLF},
			{"attribute-TAB", "attribute", whitespaceTAB},
		} {
			t.Run(tc.name+"-reference-perturbation", func(t *testing.T) {
				changed, err := perturbWhitespaceContext(b, tc.context, tc.reference)
				if err != nil {
					t.Fatal(err)
				}
				decoded, err := readWhitespace(changed)
				if err != nil {
					t.Fatal(err)
				}
				if tc.context == "text" && (decoded.Text == value || decoded.Attribute != value) {
					t.Fatalf("text perturbation not isolated: %+v", decoded)
				}
				if tc.context == "attribute" && (decoded.Attribute == value || decoded.Text != value) {
					t.Fatalf("attribute perturbation not isolated: %+v", decoded)
				}
			})
		}
	}
}

func xmlWhitespaceRoundtripSteps(sc *godog.ScenarioContext) {
	var operand, original string
	var source, input, inputBefore []byte
	var result whitespaceValue
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		operand, original = "", ""
		source, input, inputBefore = nil, nil, nil
		result = whitespaceValue{}
		return ctx, nil
	})
	sc.Step(`^an XML escaping value encoded as JSON ("x\\r\\n\\ty")$`, func(raw string) error {
		if raw != xmlWhitespaceJSON {
			return fmt.Errorf("unexpected canonical JSON input %q", raw)
		}
		input = []byte(raw)
		inputBefore = bytes.Clone(input)
		if err := json.Unmarshal(input, &operand); err != nil {
			return err
		}
		original = operand
		if operand != "x\r\n\ty" || len(operand) != 5 {
			return fmt.Errorf("wrong decoded operand %q", operand)
		}
		return nil
	})
	sc.Step(`^the value is escaped separately as text and as an attribute and both are parsed$`, func() error {
		var err error
		source, err = writeWhitespace(operand)
		if err != nil {
			return err
		}
		if _, _, err = whitespaceReferenceRegions(source); err != nil {
			return err
		}
		encodedBefore := bytes.Clone(source)
		result, err = readWhitespace(source)
		if err != nil || !bytes.Equal(source, encodedBefore) || operand != original || !bytes.Equal(input, inputBefore) {
			return fmt.Errorf("writer/reparser or source custody: %v", err)
		}
		return nil
	})
	sc.Step(`^the decoded text and attribute both equal JSON (.+)$`, func(raw string) error {
		var expected string
		if raw != xmlWhitespaceJSON {
			return fmt.Errorf("unexpected canonical JSON result %q", raw)
		}
		if err := json.Unmarshal([]byte(raw), &expected); err != nil {
			return err
		}
		if operand != original || expected != original || !bytes.Equal(input, inputBefore) || result.XMLName != (xml.Name{Local: "r"}) {
			return fmt.Errorf("source or root identity drift")
		}
		if result.Text != expected || result.Attribute != expected {
			return fmt.Errorf("decoded text=%q attribute=%q; expected %q", result.Text, result.Attribute, expected)
		}
		return nil
	})
}
