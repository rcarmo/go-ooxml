package acceptance

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/xmlsnapshot"
)

func uniformXMLCase(p *godog.Scenario) error {
	if p == nil || len(p.Steps) < 2 {
		return fmt.Errorf("missing uniform XML steps")
	}
	const prefix = "the XML source is "
	if !strings.HasPrefix(p.Steps[0].Text, prefix) {
		return fmt.Errorf("uniform XML source drift")
	}
	caller := []byte(strings.TrimPrefix(p.Steps[0].Text, prefix))
	original := bytes.Clone(caller)
	doc, err := xmlsnapshot.ParseUniform(caller)
	if err != nil {
		return err
	}
	var out []byte
	var refusal error
	wantCategory := ""
	var expected []byte
	switch {
	case hasUniformTag(p, attributeSpliceCaseID):
		if !bytes.Equal(caller, []byte(attributeSpliceSource)) {
			return fmt.Errorf("attribute source drift")
		}
		var row *struct {
			value, output string
			name          xml.Name
			attributes    [3]string
		}
		for _, name := range []string{"a", "p:n", "fresh"} {
			candidate := attributeSpliceRows[name]
			if p.Steps[1].Text == "a lexical edit sets the attribute "+name+" of the first t element to "+candidate.value {
				row = &candidate
				break
			}
		}
		if row == nil || len(doc.Elements()) != 3 {
			return fmt.Errorf("attribute edit row or source drift")
		}
		out, refusal = doc.UniformEdit(nil, []xmlsnapshot.AttributeEdit{{Target: doc.Elements()[1], Name: row.name, Value: row.value}})
		expected = []byte(row.output)
	case hasUniformTag(p, attributeRefusalCaseID):
		if !bytes.Equal(caller, []byte(attributeSpliceSource)) || len(doc.Elements()) != 3 {
			return fmt.Errorf("duplicate edit source drift")
		}
		target := doc.Elements()[1]
		out, refusal = doc.UniformEdit(nil, []xmlsnapshot.AttributeEdit{{Target: target, Name: xml.Name{Local: "a"}, Value: "x"}, {Target: target, Name: xml.Name{Local: "a"}, Value: "y"}})
		wantCategory = "overlap-or-duplicate"
	case hasUniformTag(p, childInsertionCustodyCaseID):
		if !bytes.Equal(caller, []byte(childInsertionSource)) {
			return fmt.Errorf("insertion source drift")
		}
		out, refusal = uniformChildInsertion()
		if refusal != nil {
			return refusal
		}
		// Assert the actual output of this operation independently as well.
		children := doc.Root().Children()
		if len(children) != 2 {
			return fmt.Errorf("insertion source children drift")
		}
		node := xmlsnapshot.NewElement{Name: xml.Name{Space: "new", Local: "x"}, Attributes: []xml.Attr{{Name: xml.Name{Space: "other", Local: "a"}, Value: "value"}}, Children: []xmlsnapshot.NewElement{{Name: xml.Name{Local: "plain"}, Text: "text"}}}
		another, e := doc.UniformInsert([]xmlsnapshot.ChildInsertion{{Parent: children[0], Children: []xmlsnapshot.NewElement{node}}})
		if e != nil || !bytes.Equal(out, another) {
			return fmt.Errorf("insertion repeatability: %v", e)
		}
		parsed, e := xmlsnapshot.ParseUniform(out)
		if e != nil {
			return e
		}
		var found bool
		for _, item := range parsed.Elements() {
			if item.Name() != (xml.Name{Space: "new", Local: "x"}) {
				continue
			}
			value, ok := item.Attribute("other", "a")
			kids := item.Children()
			if !ok || value != "value" || len(kids) != 1 || kids[0].Name() != (xml.Name{Local: "plain"}) {
				return fmt.Errorf("insertion authored attribute/child mismatch")
			}
			text, ok := kids[0].Text()
			if !ok || text != "text" {
				return fmt.Errorf("insertion child text mismatch")
			}
			found = true
		}
		if !found || !bytes.Contains(out, []byte(`<b>keep</b>`)) {
			return fmt.Errorf("insertion child/sibling missing")
		}
	case hasUniformTag(p, childInsertionRefusalCaseID):
		if !bytes.Equal(caller, []byte(childInsertionSource)) {
			return fmt.Errorf("insertion refusal source drift")
		}
		root := doc.Root()
		kids := root.Children()
		if len(kids) != 2 {
			return fmt.Errorf("insertion refusal children drift")
		}
		node := xmlsnapshot.NewElement{Name: xml.Name{Local: "child"}}
		out, refusal = doc.UniformInsert([]xmlsnapshot.ChildInsertion{{Parent: root, Children: []xmlsnapshot.NewElement{node}}, {Parent: kids[0], Children: []xmlsnapshot.NewElement{node}}})
		wantCategory = "overlap-or-duplicate"
	case hasUniformTag(p, elementRemovalCaseID):
		if !bytes.Equal(caller, []byte(elementRemovalSource)) {
			return fmt.Errorf("removal source drift")
		}
		all := doc.Elements()
		if len(all) != 4 {
			return fmt.Errorf("removal source nodes %d", len(all))
		}
		out, refusal = doc.UniformRemove([]xmlsnapshot.Element{all[1], all[3]})
		expected = []byte(elementRemovalResult)
	case hasUniformTag(p, elementRemovalRefusalCaseID):
		if !bytes.Equal(caller, []byte(elementRemovalSource)) {
			return fmt.Errorf("removal refusal source drift")
		}
		all := doc.Elements()
		if len(all) != 4 {
			return fmt.Errorf("removal refusal source nodes %d", len(all))
		}
		if strings.HasPrefix(p.Name, "A root ") {
			out, refusal = doc.UniformRemove([]xmlsnapshot.Element{all[0]})
			wantCategory = "root"
		} else if strings.Contains(p.Name, "p:a and its nested p:b") {
			out, refusal = doc.UniformRemove([]xmlsnapshot.Element{all[1], all[2]})
			wantCategory = "overlap-or-duplicate"
		} else {
			return fmt.Errorf("unrecognised removal refusal")
		}
	case hasUniformTag(p, elementReplacementCustodyCaseID):
		if !bytes.Equal(caller, []byte(elementReplacementSource)) {
			return fmt.Errorf("replacement source drift")
		}
		all := doc.Elements()
		if len(all) != 4 {
			return fmt.Errorf("replacement source nodes %d", len(all))
		}
		out, refusal = doc.UniformReplace([]xmlsnapshot.ElementReplacement{{Target: all[1], Nodes: []xmlsnapshot.NewElement{{Name: xml.Name{Space: "bound", Local: "new"}, Text: "value"}, {Name: xml.Name{Local: "plain"}}}}})
		expected = []byte(elementReplacementOutput)
	case hasUniformTag(p, elementReplacementRefusalCaseID):
		if !bytes.Equal(caller, []byte(elementReplacementSource)) {
			return fmt.Errorf("replacement refusal source drift")
		}
		all := doc.Elements()
		if len(all) != 4 {
			return fmt.Errorf("replacement refusal source nodes %d", len(all))
		}
		nodes := []xmlsnapshot.NewElement{{Name: xml.Name{Local: "safe"}}}
		switch {
		case strings.HasPrefix(p.Name, "A root "):
			out, refusal = doc.UniformReplace([]xmlsnapshot.ElementReplacement{{Target: all[0], Nodes: nodes}})
			wantCategory = "root"
		case strings.Contains(p.Name, "p:old twice"):
			out, refusal = doc.UniformReplace([]xmlsnapshot.ElementReplacement{{Target: all[1], Nodes: nodes}, {Target: all[1], Nodes: nodes}})
			wantCategory = "overlap-or-duplicate"
		case strings.Contains(p.Name, "p:old and its nested p:child"):
			out, refusal = doc.UniformReplace([]xmlsnapshot.ElementReplacement{{Target: all[1], Nodes: nodes}, {Target: all[2], Nodes: nodes}})
			wantCategory = "overlap-or-duplicate"
		default:
			return fmt.Errorf("unrecognised replacement refusal")
		}
	case hasUniformTag(p, immutableLeafCaseID):
		if !bytes.Equal(caller, []byte(immutableLeafSource)) {
			return fmt.Errorf("immutable source drift")
		}
		out, refusal = doc.UniformEdit(nil, nil)
		expected = []byte(immutableLeafSource)
	default:
		return fmt.Errorf("unselected uniform XML scenario %q", p.Name)
	}
	if wantCategory != "" {
		if out != nil || xmlsnapshot.UniformCategory(refusal) != wantCategory {
			return fmt.Errorf("uniform XML refusal %s: out %q err %v", wantCategory, out, refusal)
		}
	} else if refusal != nil || len(out) == 0 {
		return fmt.Errorf("uniform XML edit returned %q: %v", out, refusal)
	}
	if expected != nil && !bytes.Equal(out, expected) {
		return fmt.Errorf("uniform XML output %q, want %q", out, expected)
	}
	if !bytes.Equal(caller, original) || !bytes.Equal(doc.Source(), original) {
		return fmt.Errorf("uniform XML input/snapshot changed")
	}
	if wantCategory != "" {
		var again []byte
		var e error
		switch {
		case hasUniformTag(p, attributeRefusalCaseID), hasUniformTag(p, immutableLeafCaseID):
			again, e = doc.UniformEdit(nil, nil)
		case hasUniformTag(p, childInsertionRefusalCaseID):
			again, e = doc.UniformInsert(nil)
		case hasUniformTag(p, elementRemovalRefusalCaseID):
			again, e = doc.UniformRemove(nil)
		case hasUniformTag(p, elementReplacementRefusalCaseID):
			again, e = doc.UniformReplace(nil)
		}
		if e != nil || !bytes.Equal(again, original) {
			return fmt.Errorf("uniform XML snapshot not reusable after refusal: %v", e)
		}
	}
	// Modify the returned buffer to detect aliasing with the input and snapshot.
	if len(out) > 0 {
		out[0] = '!'
		if !bytes.Equal(caller, original) || !bytes.Equal(doc.Source(), original) {
			return fmt.Errorf("uniform XML output aliases source")
		}
	}
	return nil
}
