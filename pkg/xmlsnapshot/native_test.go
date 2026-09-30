package xmlsnapshot

import (
	"bytes"
	"encoding/xml"
	"errors"
	"strings"
	"testing"
)

func TestNativeModelOffsetsAndBounds(t *testing.T) {
	input := []byte("<?xml version=\"1.0\"?><r>😀<u:él xmlns:u=\"urn:u\" a=\"x&#13;y\"/>tail</r>")
	d, err := ParseWithLimits(input, Limits{MaxBytes: len(input), MaxSourceUnits: 100, MaxNodes: 2, MaxDepth: 2})
	if err != nil {
		t.Fatal(err)
	}
	root := d.Root()
	children := root.Children()
	if len(children) != 1 {
		t.Fatalf("children: %d", len(children))
	}
	child := children[0]
	if child.Name() != (xml.Name{Space: "urn:u", Local: "él"}) || !child.SelfClosing() {
		t.Fatalf("child: %+v", child.Name())
	}
	if parent, ok := child.Parent(); !ok || parent.Name().Local != "r" {
		t.Fatal("parent not retained")
	}
	if value, ok := child.Attribute("", "a"); !ok || value != "x\ry" {
		t.Fatalf("attribute: %q, %v", value, ok)
	}
	start := bytes.Index(input, []byte("<u:él"))
	end := bytes.Index(input, []byte("/>tail")) + 2
	rangeOfChild, err := child.SourceRange()
	if err != nil {
		t.Fatal(err)
	}
	if rangeOfChild.ByteStart != start || rangeOfChild.ByteEnd != end || rangeOfChild.UTF16Start != len([]rune(string(input[:start])))+1 {
		t.Fatalf("source range: %+v", rangeOfChild)
	}
	content, err := child.ContentRange()
	if err != nil || content.ByteStart != content.ByteEnd {
		t.Fatalf("self-close content: %+v %v", content, err)
	}
	input[0] = 'X'
	if d.Source()[0] != '<' || !bytes.Equal(childSource(d, child), []byte("<u:él xmlns:u=\"urn:u\" a=\"x&#13;y\"/>")) {
		t.Fatal("caller changed immutable source")
	}
	for name, limits := range map[string]Limits{
		"bytes":   {MaxBytes: len(input) - 1, MaxSourceUnits: 100, MaxNodes: 2, MaxDepth: 2},
		"units":   {MaxSourceUnits: 1, MaxNodes: 2, MaxDepth: 2},
		"nodes":   {MaxSourceUnits: 100, MaxNodes: 1, MaxDepth: 2},
		"depth":   {MaxSourceUnits: 100, MaxNodes: 2, MaxDepth: 1},
		"invalid": {MaxSourceUnits: 100, MaxNodes: 2, MaxDepth: 0},
	} {
		t.Run(name, func(t *testing.T) {
			got, err := ParseWithLimits(d.Source(), limits)
			if err == nil || got != nil {
				t.Fatalf("limit did not refuse: %v", err)
			}
		})
	}
}

func childSource(d *Document, e Element) []byte {
	r, _ := e.SourceRange()
	return d.Source()[r.ByteStart:r.ByteEnd]
}

func TestNativeDisjointUTF16SourceEdits(t *testing.T) {
	d, err := Parse([]byte("<r>one two</r>"))
	if err != nil {
		t.Fatal(err)
	}
	out, err := d.ApplyEdits([]SourceEdit{{Start: 7, End: 10, Value: "<x/>"}, {Start: 3, End: 6, Value: "1 &lt; 2"}})
	if err != nil || string(out) != "<r>1 &lt; 2 <x/></r>" {
		t.Fatalf("source edit: %q %v", out, err)
	}
	parsed, err := Parse(out)
	if err != nil || len(parsed.Root().Children()) != 1 {
		t.Fatalf("reparse: %v", err)
	}
	for _, tc := range []struct {
		edits    []SourceEdit
		category string
	}{
		{[]SourceEdit{{3, 5, "a"}, {4, 6, "b"}}, CategoryEditOverlap},
		{[]SourceEdit{{3, 10, "<x>"}}, CategoryEditUnsafe},
		{[]SourceEdit{{3, 10, "<!DOCTYPE x><x/>"}}, CategoryEditUnsafe},
	} {
		doc, _ := Parse([]byte("<r>text</r>"))
		got, err := doc.ApplyEdits(tc.edits)
		var editErr *EditError
		if got != nil || !errors.As(err, &editErr) || editErr.Category != tc.category {
			t.Fatalf("refusal: %q %v", got, err)
		}
	}
	unicodeDoc, _ := Parse([]byte("<r>😀ok</r>"))
	if got, err := unicodeDoc.ApplyEdits([]SourceEdit{{5, 7, "yes"}}); err != nil || string(got) != "<r>😀yes</r>" {
		t.Fatalf("UTF16 edit: %q %v", got, err)
	}
	if got, err := unicodeDoc.ApplyEdits([]SourceEdit{{4, 5, "bad"}}); got != nil || err == nil {
		t.Fatal("split surrogate accepted")
	}
}

func TestNativeEscapingAndLexicalEdits(t *testing.T) {
	for _, tc := range []struct{ input, text, attr string }{
		{"5 < 7 & 9 > 4", "5 &lt; 7 &amp; 9 &gt; 4", "5 &lt; 7 &amp; 9 &gt; 4"},
		{"'\"<&>", "'\"&lt;&amp;&gt;", "&apos;&quot;&lt;&amp;&gt;"},
		{"x\r\n\ty", "x&#13;\n\ty", "x&#13;&#10;&#9;y"},
	} {
		text, err := EscapeText(tc.input)
		if err != nil || text != tc.text {
			t.Fatalf("text %q: %q %v", tc.input, text, err)
		}
		attr, err := EscapeAttribute(tc.input)
		if err != nil || attr != tc.attr {
			t.Fatalf("attr %q: %q %v", tc.input, attr, err)
		}
		d, err := Parse([]byte("<r a=\"" + attr + "\">" + text + "</r>"))
		if err != nil {
			t.Fatal(err)
		}
		if value, _ := d.Root().Attribute("", "a"); value != tc.input {
			t.Fatalf("attribute roundtrip: %q", value)
		}
		if value, _ := d.Root().Text(); value != tc.input {
			t.Fatalf("text roundtrip: %q", value)
		}
	}
	for _, bad := range []string{"\x01", string([]byte{0xff})} {
		if output, err := EscapeText(bad); err == nil || output != "" {
			t.Fatal("invalid text escaped")
		}
		if output, err := EscapeAttribute(bad); err == nil || output != "" {
			t.Fatal("invalid attribute escaped")
		}
	}

	source := []byte("<r xmlns:p=\"u\"><!--keep--><p:t a = 'a&amp;b'/> gap <other>old</other></r>")
	d, err := Parse(source)
	if err != nil {
		t.Fatal(err)
	}
	target, other := d.Root().Children()[0], d.Root().Children()[1]
	out, err := d.Edit([]TextEdit{{Target: other, Text: "new&value"}}, []AttributeEdit{{Target: target, Name: xml.Name{Local: "a"}, Value: "x'y"}})
	if err != nil {
		t.Fatal(err)
	}
	if want := "<r xmlns:p=\"u\"><!--keep--><p:t a = 'x&#39;y'/> gap <other>new&amp;value</other></r>"; string(out) != want {
		t.Fatalf("edit result: %s", out)
	}
	if !bytes.Equal(d.Source(), source) {
		t.Fatal("edit changed snapshot")
	}
	if out, err := d.Edit(nil, nil); err != nil || !bytes.Equal(out, source) {
		t.Fatalf("no-op: %q %v", out, err)
	}
	foreign, _ := Parse(source)
	if out, err := d.Edit([]TextEdit{{Target: foreign.Root(), Text: "x"}}, nil); err == nil || out != nil {
		t.Fatal("foreign edit accepted")
	}
	if out, err := d.Edit(nil, []AttributeEdit{{Target: target, Name: xml.Name{Local: "a"}, Value: "x"}, {Target: target, Name: xml.Name{Local: "a"}, Value: "y"}}); err == nil || out != nil {
		t.Fatal("duplicate edit accepted")
	}
	removed, err := d.RemoveElements([]Element{target, other})
	if err != nil || string(removed) != "<r xmlns:p=\"u\"><!--keep--> gap </r>" {
		t.Fatalf("remove: %s %v", removed, err)
	}
	if out, err := d.RemoveElements([]Element{d.Root()}); err == nil || out != nil {
		t.Fatal("root removal accepted")
	}
	inserted, err := d.InsertChildren([]ChildInsertion{{Parent: target, Children: []NewElement{{Name: xml.Name{Space: "fresh", Local: "x"}, Text: "yes"}}}})
	if err != nil {
		t.Fatal(err)
	}
	parsed, err := Parse(inserted)
	if err != nil || parsed.Root().Children()[0].Children()[0].Name().Space != "fresh" {
		t.Fatalf("insert: %s %v", inserted, err)
	}
	replaced, err := d.ReplaceElements([]ElementReplacement{{Target: target, Nodes: []NewElement{{Name: xml.Name{Space: "u", Local: "new"}, Text: "value"}}}})
	if err != nil || !strings.Contains(string(replaced), "<p:new>value</p:new>") {
		t.Fatalf("replace: %s %v", replaced, err)
	}
	if out, err := d.ReplaceElements([]ElementReplacement{{Target: target}, {Target: target}}); err == nil || out != nil {
		t.Fatal("duplicate replacement accepted")
	}
}
