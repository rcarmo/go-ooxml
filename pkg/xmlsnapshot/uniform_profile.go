package xmlsnapshot

import (
	"bytes"
	"errors"
	"fmt"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// UniformRefusal classifies the uniform immutable XML editing profile without
// changing the existing Parse/Edit error APIs.
type UniformRefusal struct {
	Category string
	Cause    error
}

func (e *UniformRefusal) Error() string { return e.Cause.Error() }
func (e *UniformRefusal) Unwrap() error { return e.Cause }
func uniformRefuse(category string, e error) error {
	return &UniformRefusal{Category: category, Cause: e}
}
func UniformCategory(err error) string {
	var e *UniformRefusal
	if errors.As(err, &e) {
		return e.Category
	}
	return ""
}

const uniformXMLUnits = 8388608
const uniformXMLPatches = 100000

func uniformOutput(b []byte, err error) ([]byte, error) {
	if err != nil {
		return nil, uniformError(err)
	}
	if _, err := ParseWithLimits(b, Limits{MaxDepth: 256, MaxNodes: 100000, MaxSourceUnits: uniformXMLUnits}); err != nil {
		return nil, uniformError(err)
	}
	return b, nil
}
func uniformError(err error) error {
	if err == nil {
		return nil
	}
	var parse *ParseError
	if errors.As(err, &parse) {
		switch parse.Category {
		case CategoryInputTooLarge, CategoryDepthLimit, CategoryNodeLimit:
			return uniformRefuse("xml-edit-limit", err)
		case CategoryInvalidLimit:
			return uniformRefuse("invalid-name-or-value", err)
		default:
			return uniformRefuse("unsafe-XML", err)
		}
	}
	var edit *EditError
	if errors.As(err, &edit) {
		if edit.Category == CategoryEditOverlap {
			return uniformRefuse("overlap-or-duplicate", err)
		}
		return uniformRefuse("unsafe-XML", err)
	}
	var foreign *ForeignTargetError
	if errors.As(err, &foreign) {
		return uniformRefuse("foreign-target", err)
	}
	var unbound *losslessxml.UnboundAttributeNamespaceError
	if errors.As(err, &unbound) {
		return uniformRefuse("invalid-name-or-value", err)
	}
	// A novel native/programmer/I/O failure has no profile category. Do not
	// misreport it as unsafe XML merely because its diagnostic is unfamiliar.
	return err
}

// ParseUniform accepts UTF-8 XML without a leading BOM and bounded resources.
func ParseUniform(source []byte) (*Document, error) {
	if bytes.HasPrefix(source, []byte{0xef, 0xbb, 0xbf}) {
		return nil, uniformRefuse("invalid-name-or-value", fmt.Errorf("uniform XML forbids a leading BOM"))
	}
	if !utf8.Valid(source) {
		return nil, uniformRefuse("unsafe-XML", fmt.Errorf("invalid UTF-8 XML"))
	}
	d, err := ParseWithLimits(source, Limits{MaxDepth: 256, MaxNodes: 100000, MaxSourceUnits: uniformXMLUnits})
	if err != nil {
		return nil, uniformError(err)
	}
	return d, nil
}
func (d *Document) UniformEdit(texts []TextEdit, attrs []AttributeEdit) ([]byte, error) {
	if len(texts)+len(attrs) > uniformXMLPatches {
		return nil, uniformRefuse("xml-edit-limit", fmt.Errorf("XML patch budget"))
	}
	for _, edit := range texts {
		if _, err := d.owned(edit.Target); err != nil {
			return nil, uniformError(err)
		}
		if !losslessxml.ValidAuthoredText(edit.Text) {
			return nil, uniformRefuse("unsafe-XML", fmt.Errorf("invalid XML text"))
		}
	}
	for i, edit := range attrs {
		if _, err := d.owned(edit.Target); err != nil {
			return nil, uniformError(err)
		}
		if !losslessxml.ValidAuthoredName(edit.Name, true) {
			return nil, uniformRefuse("invalid-name-or-value", fmt.Errorf("invalid attribute name"))
		}
		if !losslessxml.ValidAuthoredText(edit.Value) {
			return nil, uniformRefuse("unsafe-XML", fmt.Errorf("invalid XML attribute text"))
		}
		for _, prior := range attrs[:i] {
			if edit.Target.SameElement(prior.Target) && edit.Name == prior.Name {
				return nil, uniformRefuse("overlap-or-duplicate", fmt.Errorf("duplicate attribute edit"))
			}
		}
	}
	out, err := d.Edit(texts, attrs)
	return uniformOutput(out, err)
}
func (d *Document) UniformRemove(elements []Element) ([]byte, error) {
	if len(elements) > uniformXMLPatches {
		return nil, uniformRefuse("xml-edit-limit", fmt.Errorf("XML patch budget"))
	}
	if err := d.uniformTargets(elements, true); err != nil {
		return nil, err
	}
	out, err := d.RemoveElements(elements)
	return uniformOutput(out, err)
}
func (d *Document) UniformInsert(edits []ChildInsertion) ([]byte, error) {
	if len(edits) > uniformXMLPatches {
		return nil, uniformRefuse("xml-edit-limit", fmt.Errorf("XML patch budget"))
	}
	var targets []Element
	for _, edit := range edits {
		targets = append(targets, edit.Parent)
		for _, node := range edit.Children {
			if err := uniformNode(node, 0); err != nil {
				return nil, err
			}
		}
	}
	if err := d.uniformTargets(targets, false); err != nil {
		return nil, err
	}
	out, err := d.InsertChildren(edits)
	return uniformOutput(out, err)
}
func (d *Document) UniformReplace(edits []ElementReplacement) ([]byte, error) {
	if len(edits) > uniformXMLPatches {
		return nil, uniformRefuse("xml-edit-limit", fmt.Errorf("XML patch budget"))
	}
	var targets []Element
	for _, edit := range edits {
		targets = append(targets, edit.Target)
		for _, node := range edit.Nodes {
			if err := uniformNode(node, 0); err != nil {
				return nil, err
			}
		}
	}
	if err := d.uniformTargets(targets, true); err != nil {
		return nil, err
	}
	out, err := d.ReplaceElements(edits)
	return uniformOutput(out, err)
}
func (d *Document) UniformApplyEdits(edits []SourceEdit) ([]byte, error) {
	if len(edits) > uniformXMLPatches {
		return nil, uniformRefuse("xml-edit-limit", fmt.Errorf("XML patch budget"))
	}
	out, err := d.ApplyEdits(edits)
	return uniformOutput(out, err)
}

func (d *Document) uniformTargets(targets []Element, forbidRoot bool) error {
	for i, target := range targets {
		if _, err := d.owned(target); err != nil {
			return uniformError(err)
		}
		if forbidRoot && target.SameElement(d.Root()) {
			return uniformRefuse("root", fmt.Errorf("document root edit unsupported"))
		}
		for _, prior := range targets[:i] {
			if target.SameElement(prior) {
				return uniformRefuse("overlap-or-duplicate", fmt.Errorf("duplicate element edit"))
			}
			for parent, ok := target.Parent(); ok; parent, ok = parent.Parent() {
				if parent.SameElement(prior) {
					return uniformRefuse("overlap-or-duplicate", fmt.Errorf("nested element edit"))
				}
			}
			for parent, ok := prior.Parent(); ok; parent, ok = parent.Parent() {
				if parent.SameElement(target) {
					return uniformRefuse("overlap-or-duplicate", fmt.Errorf("nested element edit"))
				}
			}
		}
	}
	return nil
}

func uniformNode(node NewElement, depth int) error {
	if depth >= 256 {
		return uniformRefuse("xml-edit-limit", fmt.Errorf("XML node depth budget"))
	}
	if !losslessxml.ValidAuthoredName(node.Name, false) || node.Text != "" && len(node.Children) != 0 {
		return uniformRefuse("invalid-name-or-value", fmt.Errorf("invalid authored XML node"))
	}
	if !losslessxml.ValidAuthoredText(node.Text) {
		return uniformRefuse("unsafe-XML", fmt.Errorf("invalid XML node text"))
	}
	seen := map[string]bool{}
	for _, attr := range node.Attributes {
		if !losslessxml.ValidAuthoredName(attr.Name, true) {
			return uniformRefuse("invalid-name-or-value", fmt.Errorf("invalid authored XML attribute"))
		}
		if !losslessxml.ValidAuthoredText(attr.Value) {
			return uniformRefuse("unsafe-XML", fmt.Errorf("invalid XML attribute text"))
		}
		key := attr.Name.Space + "\x00" + attr.Name.Local
		if seen[key] {
			return uniformRefuse("overlap-or-duplicate", fmt.Errorf("duplicate authored XML attribute"))
		}
		seen[key] = true
	}
	for _, child := range node.Children {
		if err := uniformNode(child, depth+1); err != nil {
			return err
		}
	}
	return nil
}
