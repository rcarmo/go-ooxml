package losslessxml

import (
	"bytes"
	"testing"
)

func FuzzImmutableLeafEdits(f *testing.F) {
	for _, seed := range []string{`<r><t>hello</t></r>`, `<r xmlns:p="u"><p:t a='b'>A&amp;B</p:t></r>`, `<r><t><![CDATA[a<b]]></t></r>`, `<r><t/></r>`, `<r xmlns:xml="bad"/>`} {
		f.Add(seed, "new & text")
	}
	f.Fuzz(func(t *testing.T, source, replacement string) {
		if len(source) > 65536 || len(replacement) > 4096 {
			return
		}
		input := []byte(source)
		before := bytes.Clone(input)
		d, err := Parse(input)
		if err != nil {
			return
		}
		if !bytes.Equal(input, before) {
			t.Fatal("parse mutated source")
		}
		out, err := d.Edit(nil, nil)
		if err != nil || !bytes.Equal(out, before) {
			t.Fatal("no-op not exact", err)
		}
		for _, e := range d.Elements() {
			_, leaf := e.Text()
			if !leaf {
				continue
			}
			changed, err := d.Edit([]TextEdit{{e, replacement}}, nil)
			if !bytes.Equal(input, before) {
				t.Fatal("edit mutated caller input")
			}
			if err == nil {
				if _, err = Parse(changed); err != nil {
					t.Fatal("edit produced unparseable XML", err)
				}
			}
			unchanged, err := d.Edit(nil, nil)
			if err != nil || !bytes.Equal(unchanged, before) {
				t.Fatal("edit changed snapshot", err)
			}
			break
		}
	})
}
