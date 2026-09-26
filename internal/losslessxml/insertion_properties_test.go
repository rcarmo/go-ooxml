package losslessxml

import (
	"bytes"
	"encoding/xml"
	"reflect"
	"testing"
)

// Deterministic cross-product of namespace environments, not a fuzz campaign.
func TestInsertedExpandedNamesMatrix(t *testing.T) {
	environments := []string{`<r/>`, `<r xmlns="u"/>`, `<p:r xmlns:p="u"/>`, `<r xmlns="u" xmlns:n1="v" xmlns:n2="occupied"/>`}
	spaces := []string{"", "u", "v", "fresh", "http://www.w3.org/XML/1998/namespace"}
	for _, source := range environments {
		for _, elementSpace := range spaces {
			for _, attributeSpace := range spaces {
				d, err := Parse([]byte(source))
				if err != nil {
					t.Fatal(err)
				}
				root := d.Elements()[0]
				attrs := []xml.Attr{{Name: xml.Name{Space: attributeSpace, Local: "flag"}, Value: "\t\r\n & 😀"}}
				child := NewElement{Name: xml.Name{Space: elementSpace, Local: "child"}, Attributes: attrs, Children: []NewElement{{Name: xml.Name{Local: "plain"}, Text: "x\ry\nz"}}}
				out, err := d.InsertChildren([]ChildInsertion{{Parent: root, Children: []NewElement{child}}})
				if err != nil {
					t.Fatalf("%s %q %q: %v", source, elementSpace, attributeSpace, err)
				}
				got, err := Parse(out)
				if err != nil {
					t.Fatal(err)
				}
				es := got.Elements()
				if len(es) != 3 || es[0].Name() != root.Name() || es[1].Name() != child.Name || es[2].Name() != child.Children[0].Name || !reflect.DeepEqual(es[1].Attributes(), attrs) {
					t.Fatalf("namespace meaning changed: %s", out)
				}
				text, _ := es[2].Text()
				if text != child.Children[0].Text {
					t.Fatalf("inserted text changed: %q", text)
				}
				original, err := d.Edit(nil, nil)
				if err != nil || !bytes.Equal(original, []byte(source)) {
					t.Fatal("source changed")
				}
			}
		}
	}
}

func FuzzStructuredChildRoundTrip(f *testing.F) {
	for _, pair := range [][2]string{{"a", "new"}, {"x\r\ny", "u\tv"}, {"😀&<", "\r\n\t"}} {
		f.Add(pair[0], pair[1])
	}
	f.Fuzz(func(t *testing.T, text, attribute string) {
		if len(text) > 4096 || len(attribute) > 4096 {
			return
		}
		d, err := Parse([]byte(`<p:r xmlns:p="urn:parent" xmlns:n1="occupied"/>`))
		if err != nil {
			t.Fatal(err)
		}
		node := NewElement{Name: xml.Name{Space: "urn:child", Local: "child"}, Attributes: []xml.Attr{{Name: xml.Name{Space: "urn:attribute", Local: "flag"}, Value: attribute}}, Text: text}
		out, err := d.InsertChildren([]ChildInsertion{{Parent: d.Elements()[0], Children: []NewElement{node}}})
		if !validText(text) || !validText(attribute) {
			if err == nil {
				t.Fatal("illegal characters accepted")
			}
			return
		}
		if err != nil {
			t.Fatal(err)
		}
		parsed, err := Parse(out)
		if err != nil {
			t.Fatal(err)
		}
		child := parsed.Elements()[1]
		got, _ := child.Text()
		if got != text || child.Name() != node.Name || !reflect.DeepEqual(child.Attributes(), node.Attributes) {
			t.Fatalf("authored meanings changed: %q", out)
		}
	})
}
