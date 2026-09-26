package losslessxml

import (
	"bytes"
	"testing"
)

func TestLexicalDocumentBoundaryBatch(t *testing.T) {
	for _, source := range []string{`&#32;<r/>`, `<r/>&#x20;`, `<!--a-->&#9;<r/>`, `<r/><?p ok?>&#13;`, `<![CDATA[ ]]><r/>`, `<r/><![CDATA[]]>`, `<![CDATA[]]><r/>`} {
		t.Run(source, func(t *testing.T) {
			data := []byte(source)
			if _, err := Parse(data); err == nil {
				t.Fatal("illegal boundary accepted")
			}
			if !bytes.Equal(data, []byte(source)) {
				t.Fatal("source changed")
			}
		})
	}
	for _, source := range []string{" \r\n<!--a--><?p ok?><r/>\t\n", `<r>&#32;<![CDATA[ ]]></r>`, `<?xml version="1.0"?><r a="&#32;"/>`} {
		t.Run(source, func(t *testing.T) {
			d, err := Parse([]byte(source))
			if err != nil {
				t.Fatal(err)
			}
			out, err := d.Edit(nil, nil)
			if err != nil || string(out) != source {
				t.Fatal("valid no-op changed", err)
			}
		})
	}
}
