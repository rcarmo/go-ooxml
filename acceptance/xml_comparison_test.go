package acceptance

import (
	"bytes"
	"context"
	"fmt"

	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// Keep the literal canonical operands and four-step predicate independent of the comparator.
func guardXMLComparisonCase(id string, p *messages.Pickle) error {
	type row struct{ name, left, right string }
	rows := map[string][]row{
		xmlSignificantCaseID: {
			{"XML content change to text whitespace compares different", `<a><b> x </b></a>`, `<a><b>x</b></a>`},
			{"XML content change to child order compares different", `<a><b/><c/></a>`, `<a><c/><b/></a>`},
			{"XML content change to attribute value compares different", `<a v="1"/>`, `<a v="2"/>`},
		},
		xmlPrefixBindingCaseID: {{"A prefix-valued attribute retains its namespace binding", `<a xmlns:p="urn:one" value="p:x"/>`, `<a xmlns:p="urn:two" value="p:x"/>`}},
		xmlUnsafeCaseID: {
			{"Identical DTD-bearing inputs do not compare equivalent", `<!DOCTYPE a [<!ENTITY e "text">]><a>&e;</a>`, `<!DOCTYPE a [<!ENTITY e "text">]><a>&e;</a>`},
			{"Identical malformed XML inputs do not compare equivalent", `broken`, `broken`},
		},
		xmlMarkupCaseID: {
			{"XML markup change to prolog PI target compares different", `<?one x?><a/>`, `<?two x?><a/>`},
			{"XML markup change to in-document PI target compares different", `<a><?one x?></a>`, `<a><?two x?></a>`},
			{"XML markup change to prolog comment content compares different", `<!--old--><a/>`, `<!--new--><a/>`},
		},
	}
	for _, r := range rows[id] {
		if p.Name != r.name {
			continue
		}
		expected := []string{"the left XML is " + r.left, "the right XML is " + r.right, "the conservative XML comparator compares their UTF-8 bytes", "the comparison result is false"}
		if len(p.Steps) != len(expected) {
			return fmt.Errorf("XML comparison %q step count drift", p.Name)
		}
		for i, s := range expected {
			if p.Steps[i].Text != s {
				return fmt.Errorf("XML comparison %q step %d drift", p.Name, i+1)
			}
		}
		return nil
	}
	return fmt.Errorf("unexpected XML comparison ID/row %s %q", id, p.Name)
}

func xmlComparisonSteps(sc *godog.ScenarioContext) {
	var left, right, leftCopy, rightCopy []byte
	var result bool
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		left, right, leftCopy, rightCopy = nil, nil, nil, nil
		result = false
		return ctx, nil
	})
	sc.Step(`^the left XML is (.*)$`, func(s string) error { left = []byte(s); leftCopy = bytes.Clone(left); return nil })
	sc.Step(`^the right XML is (.*)$`, func(s string) error { right = []byte(s); rightCopy = bytes.Clone(right); return nil })
	sc.Step(`^the conservative XML comparator compares their UTF-8 bytes$`, func() error {
		if left == nil || right == nil {
			return fmt.Errorf("missing XML operand")
		}
		result = losslessxml.Equivalent(left, right)
		return nil
	})
	sc.Step(`^the comparison result is false$`, func() error {
		if result || !bytes.Equal(left, leftCopy) || !bytes.Equal(right, rightCopy) {
			return fmt.Errorf("result=%t or operand changed", result)
		}
		return nil
	})
}
