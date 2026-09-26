package losslessxml

import (
	"bytes"
	"encoding/xml"
	"testing"
)

func TestStructuredChildInsertionBatch(t *testing.T) {
	d, err := Parse([]byte(`<root xmlns="u" xmlns:n1="occupied"><a/><b>keep</b></root>`))
	if err != nil {
		t.Fatal(err)
	}
	parent := d.Elements()[1]
	t.Run("fresh prefix scoped without collision", func(t *testing.T) {
		node := NewElement{Name: xml.Name{Space: "new", Local: "x"}, Attributes: []xml.Attr{{Name: xml.Name{Space: "other", Local: "a"}, Value: "v"}}, Children: []NewElement{{Name: xml.Name{Local: "plain"}, Text: "text"}}}
		b, err := d.InsertChildren([]ChildInsertion{{parent, []NewElement{node}}})
		if err != nil {
			t.Fatal(err)
		}
		p, err := Parse(b)
		if err != nil {
			t.Fatal(err)
		}
		es := p.Elements()
		if len(es) != 5 || es[2].Name() != (xml.Name{Space: "new", Local: "x"}) || es[3].Name() != (xml.Name{Local: "plain"}) {
			t.Fatalf("identities %+v", es)
		}
		if !bytes.Contains(b, []byte(`<b>keep</b>`)) {
			t.Fatal("sibling changed")
		}
	})
	t.Run("atomic invalid inputs", func(t *testing.T) {
		other, _ := Parse([]byte(`<a/>`))
		for _, edits := range [][]ChildInsertion{{{parent, []NewElement{{Name: xml.Name{Local: "bad:name"}}}}}, {{parent, []NewElement{{Name: xml.Name{Local: "x"}, Text: "\x00"}}}}, {{parent, []NewElement{{Name: xml.Name{Local: "x"}, Text: "a", Children: []NewElement{{Name: xml.Name{Local: "y"}}}}}}}, {{other.Elements()[0], []NewElement{{Name: xml.Name{Local: "x"}}}}}, {{parent, []NewElement{{Name: xml.Name{Local: "x"}}}}, {parent, []NewElement{{Name: xml.Name{Local: "y"}}}}}} {
			if _, err := d.InsertChildren(edits); err == nil {
				t.Fatal("invalid insertion accepted")
			}
		}
	})
	t.Run("nested targets and empty no op", func(t *testing.T) {
		if _, err := d.InsertChildren([]ChildInsertion{{d.Elements()[0], []NewElement{{Name: xml.Name{Local: "x"}}}}, {parent, []NewElement{{Name: xml.Name{Local: "y"}}}}}); err == nil {
			t.Fatal("overlapping insertion accepted")
		}
		out, err := d.InsertChildren(nil)
		if err != nil || !bytes.Equal(out, d.source) {
			t.Fatal("no-op changed")
		}
	})
}
