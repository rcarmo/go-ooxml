package losslessxml

import (
	"bytes"
	"encoding/xml"
	"testing"
)

func TestXMLWhitespaceNormalizationBatch(t *testing.T) {
	for _, tc := range []struct{ raw, want string }{{"\t\n\r\r\n", "    "}, {"&#9;&#10;&#13;", "\t\n\r"}, {" A  B ", " A  B "}, {"A\r\nB&#xD;\nC", "A B\r C"}, {"a&amp;b&quot;c&apos;d&lt;e&gt;", "a&b\"c'd<e>"}} {
		t.Run(tc.raw, func(t *testing.T) {
			source := []byte("<r a='" + tc.raw + "'/>")
			d, err := Parse(source)
			if err != nil {
				t.Fatal(err)
			}
			attrs := d.Elements()[0].Attributes()
			if len(attrs) != 1 || attrs[0].Value != tc.want {
				t.Fatalf("attributes %+v", attrs)
			}
			out, err := d.Edit(nil, []AttributeEdit{{d.Elements()[0], xml.Name{Local: "a"}, tc.want}})
			if err != nil || !bytes.Equal(out, source) {
				t.Fatalf("no-op changed bytes %q %v", out, err)
			}
		})
	}
	t.Run("namespace declaration values normalize", func(t *testing.T) {
		d, err := Parse([]byte("<r xmlns:p='urn:a\tb'><p:t/></r>"))
		if err != nil {
			t.Fatal(err)
		}
		if d.Elements()[1].Name().Space != "urn:a b" {
			t.Fatalf("namespace %q", d.Elements()[1].Name().Space)
		}
	})
	t.Run("CDATA line endings and original offsets", func(t *testing.T) {
		source := []byte("<r>\r\n<t><![CDATA[A\r\nB\rC]]></t>\r</r>")
		d, err := Parse(source)
		if err != nil {
			t.Fatal(err)
		}
		e := d.Elements()[1]
		text, _ := e.Text()
		if text != "A\nB\nC" {
			t.Fatalf("%q", text)
		}
		out, err := d.Edit([]TextEdit{{e, "new"}}, nil)
		if err != nil {
			t.Fatal(err)
		}
		want := []byte("<r>\r\n<t>new</t>\r</r>")
		if !bytes.Equal(out, want) {
			t.Fatalf("%q", out)
		}
	})
}
