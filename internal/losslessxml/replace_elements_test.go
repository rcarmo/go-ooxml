package losslessxml

import (
	"bytes"
	"encoding/xml"
	"testing"
)

func TestReplaceElementsBatch(t *testing.T) {
	source := []byte(`<root xmlns="outer" xmlns:p="bound"><!--a--><p:old xmlns:p="inner" x='1'><p:child/></p:old> tail <last/></root>`)
	d, err := Parse(source)
	if err != nil {
		t.Fatal(err)
	}
	old := d.Elements()[1]
	t.Run("replacement namespace is parent scope", func(t *testing.T) {
		node := NewElement{Name: xml.Name{Space: "bound", Local: "new"}, Text: "value"}
		b, err := d.ReplaceElements([]ElementReplacement{{Target: old, Nodes: []NewElement{node, {Name: xml.Name{Local: "plain"}}}}})
		if err != nil {
			t.Fatal(err)
		}
		want := `<root xmlns="outer" xmlns:p="bound"><!--a--><p:new>value</p:new><plain xmlns=""/> tail <last/></root>`
		if string(b) != want {
			t.Fatalf("%s", b)
		}
	})
	t.Run("disjoint replacement and deletion", func(t *testing.T) {
		b, err := d.ReplaceElements([]ElementReplacement{{Target: old}, {Target: d.Elements()[3], Nodes: []NewElement{{Name: xml.Name{Space: "outer", Local: "end"}}}}})
		if err != nil || string(b) != `<root xmlns="outer" xmlns:p="bound"><!--a--> tail <end/></root>` {
			t.Fatalf("%s %v", b, err)
		}
	})
	t.Run("invalid selection or content cannot escape", func(t *testing.T) {
		other, _ := Parse([]byte(`<x><y/></x>`))
		for _, edits := range [][]ElementReplacement{{{Target: d.Elements()[0]}}, {{Target: old}, {Target: old}}, {{Target: old}, {Target: d.Elements()[2]}}, {{Target: other.Elements()[1]}}, {{Target: old, Nodes: []NewElement{{Name: xml.Name{Local: "new"}, Text: "\x00"}}}}} {
			if b, err := d.ReplaceElements(edits); err == nil || b != nil {
				t.Fatal("invalid replacement exposed output")
			}
		}
		b, err := d.ReplaceElements(nil)
		if err != nil || !bytes.Equal(b, source) {
			t.Fatal("snapshot not intact")
		}
	})
}
