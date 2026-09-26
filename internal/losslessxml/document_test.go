package losslessxml

import (
	"bytes"
	"testing"
)

func TestNamespaceAndLeafContracts(t *testing.T) {
	invalid := []string{`<x:r/>`, `<r x:a="1"/>`, `<r a="1" a="2"/>`, `<r xmlns:p="u" xmlns:q="u" p:a="1" q:a="2"/>`, `<r xmlns:xml="wrong"/>`, `<r xmlns:xmlns="u"/>`, `<r xmlns:p=""/>`, `<p:r xmlns:p="u" xmlns:q="u"></q:r>`, `<r/><s/>`, `<!DOCTYPE r><r/>`, `<r xmlns="http://www.w3.org/XML/1998/namespace"/>`}
	for _, s := range invalid {
		t.Run(s, func(t *testing.T) {
			if _, err := Parse([]byte(s)); err == nil {
				t.Fatal("invalid accepted")
			}
		})
	}
	t.Run("namespace scope and attribute copies", func(t *testing.T) {
		d, err := Parse([]byte(`<r xmlns="u" xmlns:p="v"><p:t a="1" p:b="2">hello</p:t><t xmlns="">plain</t></r>`))
		if err != nil {
			t.Fatal(err)
		}
		es := d.Elements()
		if len(es) != 3 || es[1].Name().Space != "v" || es[2].Name().Space != "" {
			t.Fatalf("elements %v", es)
		}
		attrs := es[1].Attributes()
		if len(attrs) != 2 || attrs[0].Name.Space != "" || attrs[1].Name.Space != "v" {
			t.Fatalf("attrs %v", attrs)
		}
		attrs[0].Value = "changed"
		if es[1].Attributes()[0].Value != "1" {
			t.Fatal("attribute alias")
		}
	})
	t.Run("self closing remains no op but refuses insertion", func(t *testing.T) {
		src := []byte(`<r><t /></r>`)
		d, err := Parse(src)
		if err != nil {
			t.Fatal(err)
		}
		es := d.Elements()
		if len(es) != 2 {
			t.Fatal("missing elements")
		}
		out, err := d.ReplaceText([]TextEdit{{es[1], ""}})
		if err != nil || !bytes.Equal(out, src) {
			t.Fatal("no-op changes syntax")
		}
		if _, err = d.ReplaceText([]TextEdit{{es[1], "new"}}); err == nil {
			t.Fatal("self closing insertion accepted")
		}
	})
	t.Run("mixed content and comments refuse", func(t *testing.T) {
		for _, s := range []string{`<r><t>a<b/>c</t></r>`, `<r><t>a<!--marker-->b</t></r>`} {
			d, err := Parse([]byte(s))
			if err != nil {
				t.Fatal(err)
			}
			es := d.Elements()
			if len(es) < 2 {
				t.Fatal("missing elements")
			}
			if _, err = d.ReplaceText([]TextEdit{{es[1], "x"}}); err == nil {
				t.Fatal("mixed accepted")
			}
		}
	})
	t.Run("CDATA entity unicode and no op", func(t *testing.T) {
		s := []byte(`<r><t><![CDATA[a<b]]>&#x1F600;</t></r>`)
		d, err := Parse(s)
		if err != nil {
			t.Fatal(err)
		}
		es := d.Elements()
		if len(es) != 2 {
			t.Fatal("missing elements")
		}
		text, ok := es[1].Text()
		if !ok || text != "a<b😀" {
			t.Fatalf("text %q", text)
		}
		out, err := d.ReplaceText([]TextEdit{{es[1], text}})
		if err != nil || !bytes.Equal(out, s) {
			t.Fatal("no-op changed")
		}
		out, err = d.ReplaceText([]TextEdit{{es[1], "😀&"}})
		if err != nil || !bytes.Contains(out, []byte("😀&amp;")) {
			t.Fatalf("%s %v", out, err)
		}
	})
	t.Run("duplicate targets and invalid characters refuse", func(t *testing.T) {
		d, err := Parse([]byte(`<r><t>a</t></r>`))
		if err != nil {
			t.Fatal(err)
		}
		es := d.Elements()
		if len(es) != 2 {
			t.Fatal("missing elements")
		}
		for _, edits := range [][]TextEdit{{{es[1], "b"}, {es[1], "c"}}, {{es[1], "\x00"}}, {{Element{}, "bad"}}} {
			if _, err = d.ReplaceText(edits); err == nil {
				t.Fatal("invalid edits accepted")
			}
		}
	})
}
