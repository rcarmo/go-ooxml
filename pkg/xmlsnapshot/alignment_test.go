package xmlsnapshot

import (
	"errors"
	"testing"
)

func TestDirectTextExcludesDescendantsAndKeepsMarkupTails(t *testing.T) {
	d, err := ParseLexical([]byte(`<r>a<!--c-->b<?pi x?>c<x>nested</x>d<![CDATA[e]]></r>`), LexicalLimits{})
	if err != nil {
		t.Fatal(err)
	}
	root := d.Root()
	if value, _ := root.Text(); value != "abcde" {
		t.Fatalf("root direct text %q", value)
	}
	children := root.Children()
	if len(children) != 1 {
		t.Fatalf("children %d", len(children))
	}
	if value, _ := children[0].Text(); value != "nested" {
		t.Fatalf("child text %q", value)
	}
}

func TestAlignmentSourceModelFamily(t *testing.T) {
	source := []byte("<?xml version=\"1.0\"?><!--😀--><?pi ok?><p:r xmlns=\"urn:default\" xmlns:p=\"urn:p\" xmlns:x=\"urn:x\" a=\"1 &amp; 2\">pre😀<x:c x:b=\"v\"/>mid<![CDATA[<tail>]]></p:r><!--after-->")
	d, err := ParseWithLimits(source, DefaultLimits())
	if err != nil {
		t.Fatal(err)
	}
	all := d.Elements()
	if len(all) != 2 {
		t.Fatalf("element count: %d", len(all))
	}
	root, child := all[0], all[1]
	rootRange, err := root.SourceRange()
	if err != nil {
		t.Fatal(err)
	}
	childRange, err := child.SourceRange()
	if err != nil {
		t.Fatal(err)
	}
	rootContent, err := root.ContentRange()
	if err != nil {
		t.Fatal(err)
	}
	if root.QualifiedName() != "p:r" || root.Name().Space != "urn:p" || rootRange.UTF16Start != 39 || rootRange.UTF16End != 156 || rootContent.UTF16Start != 110 || rootContent.UTF16End != 150 || root.SelfClosing() {
		t.Fatalf("root: %+v %+v %s", rootRange, rootContent, root.QualifiedName())
	}
	if child.QualifiedName() != "x:c" || child.Name().Space != "urn:x" || childRange.UTF16Start != 115 || childRange.UTF16End != 129 || !child.SelfClosing() {
		t.Fatalf("child: %+v %s", childRange, child.QualifiedName())
	}
	if v, _ := root.Text(); v != "pre😀mid<tail>" {
		t.Fatalf("direct root text: %q", v)
	}
	if _, ok := root.Parent(); ok {
		t.Fatal("root parent")
	}
	if p, ok := child.Parent(); !ok || p.Name() != root.Name() || child.Root().Name() != root.Name() {
		t.Fatal("links")
	}
	if v, _ := root.Attribute("", "a"); v != "1 & 2" {
		t.Fatalf("attr: %q", v)
	}
	if v, _ := child.Attribute("urn:x", "b"); v != "v" {
		t.Fatalf("child attr: %q", v)
	}
	if string(source[rootRange.ByteStart:rootRange.ByteEnd]) != "<p:r xmlns=\"urn:default\" xmlns:p=\"urn:p\" xmlns:x=\"urn:x\" a=\"1 &amp; 2\">pre😀<x:c x:b=\"v\"/>mid<![CDATA[<tail>]]></p:r>" {
		t.Fatal("source range does not slice root")
	}
	unicode, err := ParseWithLimits([]byte("<π:名 xmlns:π=\"urn:unicode\" π:é=\"value\"><π:𐐀/></π:名>"), DefaultLimits())
	if err != nil {
		t.Fatal(err)
	}
	u, err := unicode.Root().Children()[0].SourceRange()
	if err != nil || u.UTF16Start != 39 || u.UTF16End != 46 {
		t.Fatalf("unicode: %+v %v", u, err)
	}
	line, err := ParseWithLimits([]byte("<!--😀--><r a=\"x\r\ny\tz&#xD;&#xA;&#x9;\">u\r\nv\rw&#xD;<![CDATA[c\r\nd]]><s/></r>"), DefaultLimits())
	if err != nil {
		t.Fatal(err)
	}
	if text, _ := line.Root().Text(); text != "u\nv\nw\rc\nd" {
		t.Fatalf("line text: %q", text)
	}
	if attr, _ := line.Root().Attribute("", "a"); attr != "x y z\r\n\t" {
		t.Fatalf("line attr: %q", attr)
	}
	rng, err := line.Root().Children()[0].SourceRange()
	if err != nil || rng.UTF16Start != 65 || rng.UTF16End != 69 {
		t.Fatalf("line child: %+v %v", rng, err)
	}
}

func TestBoundedDeclarationAndOffTokenClassification(t *testing.T) {
	good := []string{`<?xml version = "1.0"?><r/>`, `<?xml version="1.0" encoding = 'UTF-8' standalone = "yes"?><r/>`, `<r><!-- &custom; --><![CDATA[&custom;]]></r>`}
	for _, source := range good {
		doc, err := ParseLexical([]byte(source), LexicalLimits{})
		if err != nil || doc == nil {
			t.Errorf("valid lexical source %q: %v", source, err)
		}
	}
	bad := []struct{ source, category string }{
		{`<?xml version="1.0" version="1.0"?><r/>`, CategoryMalformedXML},
		{`<?xml encoding="UTF-8" version="1.0"?><r/>`, CategoryMalformedXML},
		{`<?xml version="1.0" standalone="yes" encoding="UTF-8"?><r/>`, CategoryMalformedXML},
		{`<r a='1'b='2'><!-- &custom; --></r>`, CategoryMalformedXML},
		{`<r a='1'b='&custom;'/>`, CategoryMalformedXML},
		{`<r a='1'b='2'><![CDATA[&custom;]]></r>`, CategoryMalformedXML},
		{`<?pi/?><r><!-- &custom; --></r>`, CategoryMalformedXML},
	}
	for _, tc := range bad {
		doc, err := ParseLexical([]byte(tc.source), LexicalLimits{})
		var typed *ParseError
		if doc != nil || !errors.As(err, &typed) || typed.Category != tc.category {
			t.Errorf("source %q: %v", tc.source, err)
		}
	}
}

func TestAlignmentRefusalAndLimitFamily(t *testing.T) {
	cases := []struct{ label, source, category string }{
		{"declaration extra", `<?xml version="1.0" extra="x"?><r/>`, CategoryMalformedXML},
		{"invalid standalone", `<?xml version="1.0" standalone="maybe"?><r/>`, CategoryMalformedXML},
		{"processing instruction", `<?pi/?><r/>`, CategoryMalformedXML},
		{"comment interior", `<r><!-- bad -- --></r>`, CategoryMalformedXML},
		{"comment termination", `<r><!--bad---></r>`, CategoryMalformedXML},
		{"attribute whitespace", `<r a='1'b='2'/>`, CategoryMalformedXML},
		{"reserved element prefix", `<xmlns:r/>`, CategoryMalformedXML},
		{"reserved default namespace", `<r xmlns="http://www.w3.org/XML/1998/namespace"/>`, CategoryMalformedXML},
		{"DTD", `<!DOCTYPE r><r/>`, CategoryDTDForbidden},
		{"undeclared entity", `<r>&custom;</r>`, CategoryEntityForbidden},
		{"lexical duplicate", `<r a='1' a='2'/>`, CategoryDuplicateAttribute},
		{"expanded duplicate", `<r xmlns:x="u" xmlns:y="u" x:a="1" y:a="2"/>`, CategoryDuplicateAttribute},
		{"unbound prefix", `<x:r/>`, CategoryUnboundPrefix},
		{"mismatched tag", `<a></b>`, CategoryMismatchedTag},
		{"invalid character", "<r>\x01</r>", CategoryInvalidCharacter},
		{"bad element local", `<p:1 xmlns:p="urn:test"/>`, CategoryMalformedXML},
		{"bad attribute local", `<r xmlns:p="urn:test" p:1="x"/>`, CategoryMalformedXML},
		{"bad prefix", `<r xmlns:1="urn:test"/>`, CategoryMalformedXML},
		{"NBSP before", "\u00a0<r/>", CategoryMalformedXML},
		{"NBSP after", "<r/>\u00a0", CategoryMalformedXML},
	}
	for _, c := range cases {
		t.Run(c.label, func(t *testing.T) {
			source := []byte(c.source)
			doc, err := ParseWithLimits(source, Limits{MaxSourceUnits: 256, MaxDepth: 4, MaxNodes: 6})
			var typed *ParseError
			if doc != nil || !errors.As(err, &typed) || typed.Category != c.category {
				t.Fatalf("result=%v err=%v category=%v want=%s", doc, err, typed, c.category)
			}
			if string(source) != c.source {
				t.Fatal("source changed")
			}
			if control, err := ParseWithLimits([]byte(`<r/>`), Limits{MaxSourceUnits: 256, MaxDepth: 4, MaxNodes: 6}); control == nil || err != nil {
				t.Fatalf("post-refusal control: %v", err)
			}
		})
	}
	for _, source := range []string{`<r><!-- <!DOCTYPE r> --></r>`, `<r><![CDATA[<!DOCTYPE r>]]></r>`} {
		if doc, err := ParseWithLimits([]byte(source), Limits{MaxSourceUnits: 256, MaxDepth: 4, MaxNodes: 6}); doc == nil || err != nil {
			t.Fatalf("false DTD refusal: %q %v", source, err)
		}
	}
	limits := Limits{MaxSourceUnits: 64, MaxDepth: 4, MaxNodes: 6}
	for _, c := range []struct{ source, category string }{
		{`<n><n><n><n><n></n></n></n></n></n>`, CategoryDepthLimit},
		{`<r><n/><n/><n/><n/><n/><n/></r>`, CategoryNodeLimit},
		{`<r>` + repeat("x", 58) + `</r>`, CategoryInputTooLarge},
	} {
		doc, err := ParseWithLimits([]byte(c.source), limits)
		var typed *ParseError
		if doc != nil || !errors.As(err, &typed) || typed.Category != c.category {
			t.Errorf("limit %q: %v", c.category, err)
		}
	}
	for _, source := range []string{`<n><n><n><n></n></n></n></n>`, `<r><n/><n/><n/><n/><n/></r>`, `<r>` + repeat("x", 57) + `</r>`} {
		if doc, err := ParseWithLimits([]byte(source), limits); doc == nil || err != nil {
			t.Errorf("positive control %q: %v", source, err)
		}
	}
}
func repeat(s string, n int) string {
	out := ""
	for range n {
		out += s
	}
	return out
}
