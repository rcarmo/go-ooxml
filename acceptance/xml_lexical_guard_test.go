package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
)

const (
	xmlParseOffsetsID    = "@id-xml-parse-offsets"
	xmlLineEndingsID     = "@id-xml-normalise-line-endings"
	xmlParseRefusalsID   = "@id-xml-parse-refusals"
	xmlParseBoundsID     = "@id-xml-parse-bounds"
	xmlInvalidQNameID    = "@id-xml-invalid-qname-components"
	xmlOutsideNBSPID     = "@id-xml-outside-root-nbsp"
	xmlEscapingValuesID  = "@id-xml-escaping-values"
	xmlEscapingInvalidID = "@id-xml-escaping-invalid-character"
	xmlApplyEditsID      = "@id-xml-apply-edits"
)

func lexicalEditingLine(defaultLine, candidateLine int) int {
	if xmlLexicalCandidate() {
		return candidateLine
	}
	return defaultLine
}

func xmlSelectedEditingCases() int {
	if xmlLexicalCandidate() {
		return 16
	}
	return 15
}
func xmlSelectedNamesCases() int {
	if xmlLexicalCandidate() {
		return 6
	}
	return 1
}

var xmlLexicalIDs = []string{xmlParseOffsetsID, xmlLineEndingsID, xmlInvalidQNameID, unicodeQNameCaseID, xmlOutsideNBSPID, xmlParseRefusalsID, xmlParseBoundsID, xmlEscapingValuesID, xmlEscapingInvalidID, xmlApplyEditsID, immutableLeafCaseID, attributeSpliceCaseID, attributeRefusalCaseID, elementRemovalCaseID, elementRemovalRefusalCaseID, childInsertionCustodyCaseID, childInsertionRefusalCaseID, childNamespaceMatrixCaseID, elementReplacementCustodyCaseID, elementReplacementRefusalCaseID}

func lexicalSelectedID(path, id string) bool {
	if !slices.Contains(xmlLexicalIDs, id) {
		return false
	}
	switch path {
	case xmlParsingFeaturePath():
		return slices.Contains([]string{xmlParseOffsetsID, xmlLineEndingsID, xmlParseRefusalsID, xmlParseBoundsID, xmlEscapingValuesID, xmlEscapingInvalidID}, id)
	case xmlNamesFeaturePath():
		return slices.Contains([]string{xmlInvalidQNameID, unicodeQNameCaseID, xmlOutsideNBSPID}, id)
	case xmlEditingFeaturePath():
		return id == xmlApplyEditsID || slices.Contains(xmlLexicalIDs[10:], id)
	}
	return false
}

// Exact candidate predicates are checked against a Go-owned fixture-shaped
// contract, not against whatever text happens to be in the current shared root.
// The sealed checkout itself is verified separately by loadReferencePin.
type lexicalShape struct {
	Name  string   `json:"name"`
	Steps []string `json:"steps"`
	Rows  []string `json:"rows"`
}

var lexicalShapes = map[string][]lexicalShape{
	xmlParseOffsetsID:    {{Name: "Parse namespaces, mixed content and preserved offsets", Steps: []string{"the lexical XML input is JSON ", "the lexical XML input is parsed without rewriting its source", "exactly two elements expose these decoded values and UTF-16 half-open offsets", "the root attribute a equals JSON \"1 & 2\" and the child expanded attribute urn:x/b equals JSON \"v\"", "the root has no parent, its sole child links back to it, and both root links identify that same root", "slicing the original source at each returned range yields its exact element markup and the source is unchanged"}}},
	xmlLineEndingsID:     {{Name: "Decode XML line endings without changing source offsets", Steps: []string{"the lexical XML input is JSON ", "the lexical XML input is parsed without rewriting its source", "the root decoded text equals JSON \"u\\nv\\nw\\rc\\nd\"", "the root attribute a equals JSON \"x y z\\r\\n\\t\"", "the child s range is UTF-16 [65,69) and slices the original source to JSON \"<s/>\"", "the original source including its raw line endings is unchanged"}}},
	xmlParseRefusalsID:   {{Name: "Refuse malformed or unsafe XML constructs", Steps: []string{"these exact XML refusal inputs and documented categories", "every refusal input is parsed through the production lexical XML API", "every input returns its documented category with no document result", "every original source remains unchanged"}}},
	xmlParseBoundsID:     {{Name: "Bound untrusted XML resources", Steps: []string{"XML parser limits maxDepth 4, maxNodes 6 and maxSourceUnits 64 measured in UTF-16 units", "these exact XML resource-limit recipes", "every recipe is parsed through the production lexical XML API with those limits", "every recipe returns its documented category with no partial document", "independent depth-four, six-node and sixty-four-source-unit controls each parse successfully", "every original source remains unchanged"}}},
	xmlInvalidQNameID:    {{Name: "Invalid element-local namespace names refuse"}, {Name: "Invalid attribute-local namespace names refuse"}, {Name: "Invalid declared-prefix namespace names refuse"}},
	xmlOutsideNBSPID:     {{Name: "A non-breaking space before the root refuses"}, {Name: "A non-breaking space after the root refuses"}},
	xmlEscapingValuesID:  {{Name: "Escape text content without changing its value"}, {Name: "Escape attribute content without changing its value"}},
	xmlEscapingInvalidID: {{Name: "Refuse an invalid XML character rather than emit it"}},
	xmlApplyEditsID:      {{Name: "Apply only disjoint edits that preserve full-document safety", Steps: []string{"the lexical XML input is JSON \"<r>one two</r>\"", "these disjoint UTF-16 half-open replacements in reverse source order", "the production XML editor applies those replacements atomically", "the complete output equals JSON \"<r>1 &lt; 2 <x/></r>\" and reparses to root text JSON \"1 < 2 \" with sole child x", "these independent edit batches refuse with no changed text", "every original source and replacement list remains unchanged"}}},
}

// Guard only new operands here. Ten published lexical snapshot IDs retain their
// existing byte/predicate guards; Unicode's old case gets a new candidate-only
// offset assertion in the live binding.
func guardLexicalCase(id string, p *messages.Pickle, line int) error {
	if id == unicodeQNameCaseID {
		if p.Name != "Unicode prefix and local names remain valid" || len(p.Steps) != 5 ||
			p.Steps[0].Text != `the lexical XML input is JSON "<π:名 xmlns:π=\"urn:unicode\" π:é=\"value\"><π:𐐀/></π:名>"` ||
			p.Steps[1].Text != "the lexical XML input is parsed without rewriting its source" ||
			p.Steps[2].Text != "the root expanded name is urn:unicode/名 and its expanded attribute urn:unicode/é equals JSON \"value\"" ||
			p.Steps[3].Text != "its sole child expanded name is urn:unicode/𐐀 with UTF-16 range [39,46) slicing to JSON \"<π:𐐀/>\"" ||
			p.Steps[4].Text != "the root UTF-16 range is [0,52) and the original source is unchanged" {
			return fmt.Errorf("Unicode QName lexical candidate drift")
		}
		return nil
	}
	shapes := lexicalShapes[id]
	if len(shapes) == 0 {
		return fmt.Errorf("unknown lexical ID %s", id)
	}
	var want *lexicalShape
	for i := range shapes {
		if shapes[i].Name == p.Name {
			want = &shapes[i]
			break
		}
	}
	if want == nil || line <= 0 || len(p.AstNodeIds) == 0 {
		return fmt.Errorf("lexical case %s name/row drift: %q", id, p.Name)
	}
	if len(want.Steps) > 0 {
		if len(p.Steps) != len(want.Steps) {
			return fmt.Errorf("lexical steps %s count drift", id)
		}
		for i, step := range p.Steps {
			prefix := want.Steps[i]
			if step.Text != prefix && !(i == 0 && (id == xmlParseOffsetsID || id == xmlLineEndingsID) && len(step.Text) > len(prefix) && step.Text[:len(prefix)] == prefix) {
				return fmt.Errorf("lexical step %s #%d drift: %q", id, i+1, step.Text)
			}
		}
	}
	return nil
}

func guardLexicalFeature(doc *messages.GherkinDocument, path string) error {
	if doc == nil || doc.Feature == nil {
		return fmt.Errorf("missing lexical feature")
	}
	expected := map[string]int{}
	switch path {
	case xmlParsingFeaturePath():
		for _, id := range []string{xmlParseOffsetsID, xmlLineEndingsID, xmlParseRefusalsID, xmlParseBoundsID, xmlEscapingValuesID, xmlEscapingInvalidID} {
			expected[id] = 1
		}
	case xmlNamesFeaturePath():
		expected[xmlInvalidQNameID] = 1
		expected[xmlOutsideNBSPID] = 1
		expected[unicodeQNameCaseID] = 1
	case xmlEditingFeaturePath():
		expected[xmlApplyEditsID] = 1
	default:
		return fmt.Errorf("unexpected lexical path")
	}
	seen := map[string]int{}
	visit := func(s *messages.Scenario) {
		if s == nil {
			return
		}
		for _, tag := range s.Tags {
			if _, ok := expected[tag.Name]; ok {
				seen[tag.Name]++
			}
		}
	}
	for _, child := range doc.Feature.Children {
		visit(child.Scenario)
		if child.Rule != nil {
			for _, member := range child.Rule.Children {
				visit(member.Scenario)
			}
		}
	}
	for id, count := range expected {
		if seen[id] != count {
			return fmt.Errorf("lexical ID %s count %d not %d", id, seen[id], count)
		}
	}
	return nil
}

func TestLexicalCandidateGuardRejectsDrift(t *testing.T) {
	loadReferencePin(t)
	if !xmlLexicalCandidate() {
		t.Skip("requires sealed XML lexical candidate")
	}
	for _, path := range []string{xmlParsingFeaturePath(), xmlNamesFeaturePath(), xmlEditingFeaturePath()} {
		t.Run(filepath.Base(path), func(t *testing.T) {
			file, err := os.Open(path)
			if err != nil {
				t.Fatal(err)
			}
			n := 0
			next := func() string { n++; return fmt.Sprint(n) }
			doc, err := gherkin.ParseGherkinDocument(file, next)
			_ = file.Close()
			if err != nil {
				t.Fatal(err)
			}
			if err := guardLexicalFeature(doc, path); err != nil {
				t.Fatal(err)
			}
			for _, p := range gherkin.Pickles(*doc, path, next) {
				for _, tag := range p.Tags {
					if lexicalSelectedID(path, tag.Name) && len(lexicalShapes[tag.Name]) > 0 {
						if err := guardLexicalCase(tag.Name, p, 1); err != nil {
							t.Fatal(err)
						}
						clone := *p
						clone.Name += " changed"
						if err := guardLexicalCase(tag.Name, &clone, 1); err == nil {
							t.Fatalf("guard accepted name drift: %s", tag.Name)
						}
					}
				}
			}
		})
	}
}

func decodeLexicalJSON(raw string) (string, error) {
	var s string
	if err := json.Unmarshal([]byte(raw), &s); err != nil {
		return "", err
	}
	return s, nil
}
