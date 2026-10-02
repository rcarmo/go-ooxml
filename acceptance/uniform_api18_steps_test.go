package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"reflect"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/staticref"
	"github.com/rcarmo/go-ooxml/pkg/xmlsnapshot"
)

// The historical steps remain bound to the legacy APIs. These additional
// predicates execute the same sealed case through the public uniform APIs.
// Expected records and bytes come from the separately sealed feature layout.
func uniformAPI18Steps(sc *godog.ScenarioContext) {
	var scenario *godog.Scenario
	var source string
	var refs []staticref.Reference
	var rangeValue staticref.Range
	var failure error
	var evaluated bool
	sc.Before(func(ctx context.Context, p *godog.Scenario) (context.Context, error) {
		scenario, source, refs, rangeValue, failure, evaluated = p, "", nil, staticref.Range{}, nil, false
		return ctx, nil
	})
	formulaSource := func() error {
		if scenario == nil || len(scenario.Steps) < 2 {
			return fmt.Errorf("missing uniform formula case")
		}
		first := scenario.Steps[0].Text
		const prefix = "the formula source is JSON "
		if !strings.HasPrefix(first, prefix) {
			return fmt.Errorf("uniform formula source drift: %q", first)
		}
		encoded := strings.TrimSuffix(strings.TrimPrefix(first, prefix), " in context sheet Main")
		if err := json.Unmarshal([]byte(encoded), &source); err != nil {
			return err
		}
		return nil
	}
	sc.Step(`^the complete normalized reference records( or refusal)? equal JSON (.+)$`, func(_, encoded string) error {
		if err := formulaSource(); err != nil {
			return err
		}
		refs, failure = staticref.Analyze(source)
		evaluated = true
		return compareUniformReferences(source, refs, failure, encoded)
	})
	sc.Step(`^the complete normalized direct range equals JSON (.+)$`, func(encoded string) error {
		if scenario == nil || len(scenario.Steps) < 2 {
			return fmt.Errorf("missing uniform direct range")
		}
		const prefix = "the direct range source is JSON "
		if !strings.HasPrefix(scenario.Steps[0].Text, prefix) {
			return fmt.Errorf("direct range source drift")
		}
		if err := json.Unmarshal([]byte(strings.TrimPrefix(scenario.Steps[0].Text, prefix)), &source); err != nil {
			return err
		}
		rangeValue, failure = staticref.ParseRange(source)
		evaluated = true
		var expected uniformRange
		if err := json.Unmarshal([]byte(encoded), &expected); err != nil {
			return err
		}
		actual := uniformRange{Sheet: rangeValue.Sheet, Axis: rangeValue.Axis, First: toUniformCell(rangeValue.First), Last: toUniformCell(rangeValue.Last)}
		if failure != nil || actual != expected {
			return fmt.Errorf("direct range %q: got %+v, error %v; want %+v", source, actual, failure, expected)
		}
		return nil
	})
	sc.Step(`^the authored attribute expanded name is other/a with JSON value "value" and the plain grandchild text is JSON "text"$`, func() error {
		if scenario == nil || !hasUniformTag(scenario, childInsertionCustodyCaseID) {
			return fmt.Errorf("unselected authored-attribute assertion")
		}
		out, err := uniformChildInsertion()
		if err != nil {
			return err
		}
		doc, err := xmlsnapshot.ParseUniform(out)
		if err != nil {
			return err
		}
		for _, node := range doc.Elements() {
			if node.Name() != (xml.Name{Space: "new", Local: "x"}) {
				continue
			}
			if v, ok := node.Attribute("other", "a"); !ok || v != "value" {
				return fmt.Errorf("authored attribute: %q %v", v, ok)
			}
			children := node.Children()
			if len(children) != 1 || children[0].Name() != (xml.Name{Local: "plain"}) {
				return fmt.Errorf("plain child: %+v", children)
			}
			if v, ok := children[0].Text(); !ok || v != "text" {
				return fmt.Errorf("authored grandchild text: %q %v", v, ok)
			}
			return nil
		}
		return fmt.Errorf("inserted new/x missing")
	})
	sc.Step(`^the uniform profile result, refusal category and immutable input custody match the sealed API contract$`, func() error {
		if scenario == nil {
			return fmt.Errorf("uniform case missing")
		}
		switch {
		case hasUniformTag(scenario, formulaAnalysisCountsCaseID), hasUniformTag(scenario, formulaQuotedSheetFlagsCaseID), hasUniformTag(scenario, formulaAnalysisRefusalCaseID), hasUniformTag(scenario, formulaLiteralPunctuationCaseID):
			if !evaluated {
				if err := formulaSource(); err != nil {
					return err
				}
				refs, failure = staticref.Analyze(source)
			}
			if failure != nil && (refs != nil || staticCategory(failure) != "unsupported-static-reference") {
				return fmt.Errorf("formula refusal %q: refs %+v err %v", source, refs, failure)
			}
			if failure == nil && refs == nil {
				return fmt.Errorf("formula %q returned nil references", source)
			}
			for _, r := range refs {
				if r.Start < 0 || r.End <= r.Start || r.End > len(source) {
					return fmt.Errorf("formula span %+v", r)
				}
				one, err := staticref.Analyze(source[r.Start:r.End])
				if err != nil || len(one) != 1 || one[0].Start != 0 || one[0].End != r.End-r.Start || one[0].First != r.First || one[0].Last != r.Last || one[0].Sheet != r.Sheet {
					return fmt.Errorf("independent reference slice %q: %+v %v", source[r.Start:r.End], one, err)
				}
			}
			return nil
		case hasUniformTag(scenario, directRangeParsingCaseID), hasUniformTag(scenario, directRangeRefusalCaseID):
			if !evaluated {
				const prefix = "the direct range source is JSON "
				if !strings.HasPrefix(scenario.Steps[0].Text, prefix) {
					return fmt.Errorf("direct source drift")
				}
				if err := json.Unmarshal([]byte(strings.TrimPrefix(scenario.Steps[0].Text, prefix)), &source); err != nil {
					return err
				}
				rangeValue, failure = staticref.ParseRange(source)
			}
			if hasUniformTag(scenario, directRangeRefusalCaseID) {
				if rangeValue != (staticref.Range{}) || staticCategory(failure) != "unsupported-direct-range" {
					return fmt.Errorf("direct refusal %q: %+v %v", source, rangeValue, failure)
				}
			} else if failure != nil {
				return fmt.Errorf("direct success %q: %v", source, failure)
			}
			return nil
		case hasUniformTag(scenario, staticRemapExactCaseID), hasUniformTag(scenario, staticRemapRefusalCaseID):
			return uniformRemapCase(scenario)
		case hasUniformTag(scenario, staticReferencePropertiesCaseID):
			return uniformMatrixCase()
		default:
			return uniformXMLCase(scenario)
		}
	})
}

func hasUniformTag(p *godog.Scenario, id string) bool {
	for _, tag := range p.Tags {
		if tag.Name == id {
			return true
		}
	}
	return false
}
func staticCategory(err error) string {
	if e, ok := err.(*staticref.Refusal); ok {
		return e.Category
	}
	return ""
}

type uniformCell struct {
	Row            int  `json:"row"`
	Column         int  `json:"column"`
	RowAbsolute    bool `json:"rowAbsolute"`
	ColumnAbsolute bool `json:"columnAbsolute"`
}
type uniformReference struct {
	Sheet string      `json:"sheet"`
	Axis  string      `json:"axis"`
	First uniformCell `json:"first"`
	Last  uniformCell `json:"last"`
	Start int         `json:"start"`
	End   int         `json:"end"`
}
type uniformRange struct {
	Sheet string      `json:"sheet"`
	Axis  string      `json:"axis"`
	First uniformCell `json:"first"`
	Last  uniformCell `json:"last"`
}

func toUniformCell(c staticref.Cell) uniformCell {
	return uniformCell{c.Row, c.Column, c.RowAbsolute, c.ColumnAbsolute}
}
func compareUniformReferences(source string, refs []staticref.Reference, refusal error, encoded string) error {
	if encoded == "null" {
		if refs != nil || staticCategory(refusal) != "unsupported-static-reference" {
			return fmt.Errorf("expected atomic formula refusal for %q: %+v %v", source, refs, refusal)
		}
		return nil
	}
	var expected []uniformReference
	if err := json.Unmarshal([]byte(encoded), &expected); err != nil {
		return err
	}
	actual := make([]uniformReference, 0, len(refs))
	for _, r := range refs {
		actual = append(actual, uniformReference{r.Sheet, r.Axis, toUniformCell(r.First), toUniformCell(r.Last), r.Start, r.End})
	}
	if refusal != nil || !reflect.DeepEqual(actual, expected) {
		return fmt.Errorf("formula %q: got %+v, error %v; want %+v", source, actual, refusal, expected)
	}
	return nil
}
func uniformRemapCase(p *godog.Scenario) error {
	if len(p.Steps) < 3 {
		return fmt.Errorf("remap steps missing")
	}
	const sourcePrefix = "the formula source is JSON "
	if !strings.HasPrefix(p.Steps[0].Text, sourcePrefix) || !strings.HasSuffix(p.Steps[0].Text, " in context sheet Main") {
		return fmt.Errorf("remap source drift")
	}
	var source, sheet string
	s := strings.TrimSuffix(strings.TrimPrefix(p.Steps[0].Text, sourcePrefix), " in context sheet Main")
	if err := json.Unmarshal([]byte(s), &source); err != nil {
		return err
	}
	var axis string
	var at, count int
	var encoded string
	if _, err := fmt.Sscanf(p.Steps[1].Text, "the static remapper inserts %s at %d by %d on sheet JSON %s", &axis, &at, &count, &encoded); err != nil {
		return err
	}
	if err := json.Unmarshal([]byte(encoded), &sheet); err != nil {
		return err
	}
	out, failure := staticref.InsertReferences(source, "Main", staticref.Insertion{Sheet: sheet, Axis: axis, At: at, Count: count})
	if hasUniformTag(p, staticRemapRefusalCaseID) {
		category := "invalid-reference-insertion"
		if strings.Contains(source, "INDIRECT(") {
			category = "unsupported-static-reference"
		}
		if out != "" || staticCategory(failure) != category {
			return fmt.Errorf("remap refusal %q: output %q error %v expected %s", source, out, failure, category)
		}
		return nil
	}
	const expectedPrefix = "the complete replacement expression equals JSON "
	if !strings.HasPrefix(p.Steps[2].Text, expectedPrefix) {
		return fmt.Errorf("remap output drift")
	}
	var expected string
	if err := json.Unmarshal([]byte(strings.TrimPrefix(p.Steps[2].Text, expectedPrefix)), &expected); err != nil {
		return err
	}
	if failure != nil || out != expected {
		return fmt.Errorf("remap %q: %q %v want %q", source, out, failure, expected)
	}
	if unchanged, err := staticref.Analyze(source); err != nil || len(unchanged) == 0 {
		return fmt.Errorf("remap original source changed: %+v %v", unchanged, err)
	}
	return nil
}
func uniformMatrixCase() error {
	matrix := makeFormulaMatrix()
	if len(matrix) != 288 {
		return fmt.Errorf("matrix size %d", len(matrix))
	}
	for i, entry := range matrix {
		refs, err := staticref.Analyze(entry.source)
		if err != nil || len(refs) != 2 {
			return fmt.Errorf("uniform matrix %d refs %+v err %v", i, refs, err)
		}
		for n, r := range refs {
			want := entry.expected[n]
			if r.Sheet != want.Sheet || r.Axis != "cell" || r.First != (staticref.Cell{Row: want.First.Row, Column: want.First.Column, RowAbsolute: want.First.AbsoluteRow, ColumnAbsolute: want.First.AbsoluteColumn}) || r.Last != (staticref.Cell{Row: want.Last.Row, Column: want.Last.Column, RowAbsolute: want.Last.AbsoluteRow, ColumnAbsolute: want.Last.AbsoluteColumn}) || r.Start != want.Start || r.End != want.End || entry.source[r.Start:r.End] != entry.raw[n] {
				return fmt.Errorf("uniform matrix %d ref %d: %+v", i, n, r)
			}
			parsed, parseErr := staticref.Analyze(entry.raw[n])
			if parseErr != nil || len(parsed) != 1 || parsed[0].Start != 0 || parsed[0].End != len(entry.raw[n]) || parsed[0].Sheet != r.Sheet || parsed[0].First != r.First || parsed[0].Last != r.Last {
				return fmt.Errorf("uniform matrix %d reparse %q: %+v %v", i, entry.raw[n], parsed, parseErr)
			}
		}
		out, err := staticref.InsertReferences(entry.source, "Main", staticref.Insertion{Sheet: "Main", Axis: "row", At: 100, Count: 1})
		if err != nil || out != entry.source {
			return fmt.Errorf("uniform matrix %d no-op %q %v", i, out, err)
		}
		wrapped, err := staticref.Analyze("SUM(" + entry.source + ")")
		if err != nil || len(wrapped) != 2 {
			return fmt.Errorf("uniform matrix %d wrap %+v %v", i, wrapped, err)
		}
		for n, ref := range wrapped {
			want := refs[n]
			want.Start += len("SUM(")
			want.End += len("SUM(")
			if ref != want || ("SUM(" + entry.source + ")")[ref.Start:ref.End] != entry.raw[n] {
				return fmt.Errorf("uniform matrix %d wrapped ref %d: %+v want %+v", i, n, ref, want)
			}
		}
	}
	return nil
}

func uniformChildInsertion() ([]byte, error) {
	caller := []byte(childInsertionSource)
	doc, err := xmlsnapshot.ParseUniform(caller)
	if err != nil {
		return nil, err
	}
	root := doc.Root()
	children := root.Children()
	if len(children) != 2 {
		return nil, fmt.Errorf("source children %d", len(children))
	}
	node := xmlsnapshot.NewElement{Name: xml.Name{Space: "new", Local: "x"}, Attributes: []xml.Attr{{Name: xml.Name{Space: "other", Local: "a"}, Value: "value"}}, Children: []xmlsnapshot.NewElement{{Name: xml.Name{Local: "plain"}, Text: "text"}}}
	out, err := doc.UniformInsert([]xmlsnapshot.ChildInsertion{{Parent: children[0], Children: []xmlsnapshot.NewElement{node}}})
	if err != nil {
		return nil, err
	}
	if !bytes.Equal(caller, []byte(childInsertionSource)) || !bytes.Equal(doc.Source(), caller) || !bytes.Contains(out, []byte(`<b>keep</b>`)) {
		return nil, fmt.Errorf("insertion input/sibling custody failed")
	}
	return out, nil
}
