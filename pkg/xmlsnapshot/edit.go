package xmlsnapshot

import (
	"encoding/xml"
	"fmt"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// TextEdit replaces the decoded text of a plain leaf in its originating snapshot.
type TextEdit struct {
	Target Element
	Text   string
}

// AttributeEdit sets a non-namespace attribute by expanded name.
type AttributeEdit struct {
	Target Element
	Name   xml.Name
	Value  string
}

// NewElement is a structured replacement or insertion, never raw XML.
type NewElement = losslessxml.NewElement

type ChildInsertion struct {
	Parent   Element
	Children []NewElement
}

type ElementReplacement struct {
	Target Element
	Nodes  []NewElement
}

func (d *Document) owned(e Element) (losslessxml.Element, error) {
	if d == nil || d.doc == nil || e.doc != d.doc || e.element.Ordinal() < 0 {
		return losslessxml.Element{}, fmt.Errorf("foreign or invalid XML snapshot target")
	}
	return e.element, nil
}

// Edit atomically changes text and attributes while preserving all untouched
// source bytes. It returns no output for a refused batch.
func (d *Document) Edit(texts []TextEdit, attributes []AttributeEdit) ([]byte, error) {
	if d == nil || d.doc == nil {
		return nil, fmt.Errorf("nil XML snapshot")
	}
	t := make([]losslessxml.TextEdit, 0, len(texts))
	a := make([]losslessxml.AttributeEdit, 0, len(attributes))
	for _, edit := range texts {
		target, err := d.owned(edit.Target)
		if err != nil {
			return nil, err
		}
		t = append(t, losslessxml.TextEdit{Target: target, Text: edit.Text})
	}
	for _, edit := range attributes {
		target, err := d.owned(edit.Target)
		if err != nil {
			return nil, err
		}
		a = append(a, losslessxml.AttributeEdit{Target: target, Name: edit.Name, Value: edit.Value})
	}
	return d.doc.Edit(t, a)
}

// RemoveElements removes disjoint non-root subtrees from one snapshot.
func (d *Document) RemoveElements(elements []Element) ([]byte, error) {
	if d == nil || d.doc == nil {
		return nil, fmt.Errorf("nil XML snapshot")
	}
	targets := make([]losslessxml.Element, 0, len(elements))
	for _, e := range elements {
		target, err := d.owned(e)
		if err != nil {
			return nil, err
		}
		targets = append(targets, target)
	}
	return d.doc.RemoveElements(targets)
}

// InsertChildren appends structured nodes under disjoint owned parents.
func (d *Document) InsertChildren(insertions []ChildInsertion) ([]byte, error) {
	if d == nil || d.doc == nil {
		return nil, fmt.Errorf("nil XML snapshot")
	}
	items := make([]losslessxml.ChildInsertion, 0, len(insertions))
	for _, insertion := range insertions {
		parent, err := d.owned(insertion.Parent)
		if err != nil {
			return nil, err
		}
		items = append(items, losslessxml.ChildInsertion{Parent: parent, Children: insertion.Children})
	}
	return d.doc.InsertChildren(items)
}

// ReplaceElements swaps disjoint non-root subtrees in their surviving parent scope.
func (d *Document) ReplaceElements(replacements []ElementReplacement) ([]byte, error) {
	if d == nil || d.doc == nil {
		return nil, fmt.Errorf("nil XML snapshot")
	}
	items := make([]losslessxml.ElementReplacement, 0, len(replacements))
	for _, replacement := range replacements {
		target, err := d.owned(replacement.Target)
		if err != nil {
			return nil, err
		}
		items = append(items, losslessxml.ElementReplacement{Target: target, Nodes: replacement.Nodes})
	}
	return d.doc.ReplaceElements(items)
}
