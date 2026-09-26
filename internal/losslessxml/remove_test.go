package losslessxml

import (
	"bytes"
	"testing"
)

func TestElementRemovalBatch(t *testing.T) {
	source := []byte(`<r xmlns:p="u"><!--keep--><p:a x = '1'><p:b>text</p:b></p:a> gap <p:c /></r>`)
	d, err := Parse(source)
	if err != nil {
		t.Fatal(err)
	}
	es := d.Elements()
	out, err := d.RemoveElements([]Element{es[1], es[3]})
	if err != nil || string(out) != `<r xmlns:p="u"><!--keep--> gap </r>` {
		t.Fatalf("%s %v", out, err)
	}
	foreign, _ := Parse([]byte(`<x/>`))
	for _, targets := range [][]Element{{es[0]}, {es[1], es[1]}, {es[1], es[2]}, {foreign.Elements()[0]}, {Element{}}} {
		if _, err = d.RemoveElements(targets); err == nil {
			t.Fatal("invalid removal accepted")
		}
	}
	out, err = d.RemoveElements(nil)
	if err != nil || !bytes.Equal(source, out) {
		t.Fatal("no-op changed")
	}
	if _, err = d.RemoveElements([]Element{es[2]}); err != nil {
		t.Fatal("snapshot not reusable", err)
	}
}
