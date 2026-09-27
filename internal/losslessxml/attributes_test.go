package losslessxml

import (
	"encoding/xml"
	"testing"
)

func TestAttributeSpliceBatch(t *testing.T) {
	d, err := Parse([]byte(`<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="old"/><t>text</t></r>`))
	if err != nil {
		t.Fatal(err)
	}
	e := d.Elements()[1]
	for _, tc := range []struct {
		name        xml.Name
		value, want string
	}{{xml.Name{Local: "a"}, "x'y&z", `<r xmlns:p="urn:p"><t a = 'x&#39;y&amp;z' p:n="old"/><t>text</t></r>`}, {xml.Name{Space: "urn:p", Local: "n"}, "new", `<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="new"/><t>text</t></r>`}, {xml.Name{Local: "fresh"}, "value", `<r xmlns:p="urn:p"><t a = 'a&amp;b' p:n="old" fresh="value"/><t>text</t></r>`}} {
		t.Run(tc.value, func(t *testing.T) {
			got, err := d.Edit(nil, []AttributeEdit{{e, tc.name, tc.value}})
			if err != nil || string(got) != tc.want {
				t.Fatalf("%s %v", got, err)
			}
		})
	}
	for _, edits := range [][]AttributeEdit{{{e, xml.Name{Space: "urn:absent", Local: "x"}, "v"}}, {{e, xml.Name{Local: "xmlns"}, "v"}}, {{e, xml.Name{Local: "a"}, "x"}, {e, xml.Name{Local: "a"}, "y"}}, {{e, xml.Name{Local: "a"}, "\x00"}}, {{e, xml.Name{Local: "bad:name"}, "v"}}} {
		if _, err := d.Edit(nil, edits); err == nil {
			t.Fatal("invalid attribute batch accepted")
		}
	}
}
