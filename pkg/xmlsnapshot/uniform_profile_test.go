package xmlsnapshot

import (
	"bytes"
	"encoding/xml"
	"errors"
	"strings"
	"testing"
)

func TestUniformXMLProfileRedBatch(t *testing.T) {
	source := []byte(`<r><x>😀</x><y/></r>`)
	doc, err := ParseUniform(source)
	if err != nil {
		t.Fatal(err)
	}
	root := doc.Root()
	elements := doc.Elements()
	if len(elements) != 3 {
		t.Fatal("elements")
	}
	if out, err := doc.UniformRemove([]Element{root}); out != nil || UniformCategory(err) != "root" {
		t.Fatalf("root %q %v", out, err)
	}
	other, err := ParseUniform(source)
	if err != nil {
		t.Fatal(err)
	}
	if out, err := doc.UniformRemove([]Element{other.Elements()[1]}); out != nil || UniformCategory(err) != "foreign-target" {
		t.Fatalf("foreign %q %v", out, err)
	}
	if out, err := doc.UniformRemove([]Element{elements[1], elements[1]}); out != nil || UniformCategory(err) != "overlap-or-duplicate" {
		t.Fatalf("duplicate %q %v", out, err)
	}
	if out, err := doc.UniformInsert([]ChildInsertion{{Parent: root, Children: []NewElement{{Name: xml.Name{Local: "z"}, Text: "v"}}}}); err != nil || bytes.Equal(out, source) {
		t.Fatalf("insert %q %v", out, err)
	}
	if out, err := doc.UniformApplyEdits([]SourceEdit{{Start: 7, End: 8, Value: "a"}}); out != nil || UniformCategory(err) != "unsafe-XML" {
		t.Fatalf("split %q %v", out, err)
	}
	if out, err := doc.UniformEdit(nil, nil); err != nil || !bytes.Equal(out, source) {
		t.Fatalf("noop %q %v", out, err)
	}
	if _, err := ParseUniform([]byte("\xef\xbb\xbf<r/>")); UniformCategory(err) != "invalid-name-or-value" {
		t.Fatalf("BOM %v", err)
	}
}

// These outputs begin from a valid, tiny snapshot. ApplyEdits validates its
// result internally and wraps a typed parser refusal; the uniform facade must
// preserve the cause and report a limit, never generic unsafe XML.
func TestUniformXMLOutputLimitCauses(t *testing.T) {
	source := []byte(`<r></r>`)
	doc, err := ParseUniform(source)
	if err != nil {
		t.Fatal(err)
	}
	cases := []struct{ name, fragment, native string }{
		{"depth", strings.Repeat("<a>", 256) + strings.Repeat("</a>", 256), CategoryDepthLimit},
		{"nodes", strings.Repeat("<a/>", 100000), CategoryNodeLimit},
		{"units", strings.Repeat("a", uniformXMLUnits-6), CategoryInputTooLarge},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			out, refusal := doc.UniformApplyEdits([]SourceEdit{{Start: 3, End: 3, Value: tc.fragment}})
			var native *ParseError
			if out != nil || UniformCategory(refusal) != "xml-edit-limit" || !errors.As(refusal, &native) || native.Category != tc.native {
				t.Fatalf("out len %d category %q typed %+v err %v", len(out), UniformCategory(refusal), native, refusal)
			}
			if !bytes.Equal(source, []byte(`<r></r>`)) || !bytes.Equal(doc.Source(), source) {
				t.Fatal("source or snapshot changed")
			}
			unchanged, err := doc.UniformApplyEdits(nil)
			if err != nil || !bytes.Equal(unchanged, source) {
				t.Fatalf("snapshot not reusable: %v", err)
			}
		})
	}
	// Legacy structured/text editors reparse without uniform limits. Their
	// successful bytes must be checked by the public profile before return.
	t.Run("legacy text output units", func(t *testing.T) {
		leaf, err := ParseUniform([]byte(`<r><t>x</t></r>`))
		if err != nil {
			t.Fatal(err)
		}
		original := leaf.Source()
		out, refusal := leaf.UniformEdit([]TextEdit{{Target: leaf.Elements()[1], Text: strings.Repeat("z", uniformXMLUnits)}}, nil)
		var native *ParseError
		if out != nil || UniformCategory(refusal) != "xml-edit-limit" || !errors.As(refusal, &native) || native.Category != CategoryInputTooLarge || !bytes.Equal(leaf.Source(), original) {
			t.Fatalf("text output %d category %q cause %+v err %v", len(out), UniformCategory(refusal), native, refusal)
		}
	})
	t.Run("legacy structured output depth", func(t *testing.T) {
		leaf, err := ParseUniform([]byte(`<r><a/></r>`))
		if err != nil {
			t.Fatal(err)
		}
		original := leaf.Source()
		node := NewElement{Name: xml.Name{Local: "n"}}
		for i := 0; i < 255; i++ {
			node = NewElement{Name: xml.Name{Local: "n"}, Children: []NewElement{node}}
		}
		out, refusal := leaf.UniformReplace([]ElementReplacement{{Target: leaf.Elements()[1], Nodes: []NewElement{node}}})
		var native *ParseError
		if out != nil || UniformCategory(refusal) != "xml-edit-limit" || !errors.As(refusal, &native) || native.Category != CategoryDepthLimit || !bytes.Equal(leaf.Source(), original) {
			t.Fatalf("structured output %d category %q cause %+v", len(out), UniformCategory(refusal), native)
		}
	})
}

func TestUniformXMLBoundaryBatch(t *testing.T) {
	source := []byte(`<r><a/><b/></r>`)
	one, err := ParseUniform(source)
	if err != nil {
		t.Fatal(err)
	}
	two, err := ParseUniform(source)
	if err != nil {
		t.Fatal(err)
	}
	for _, test := range []struct {
		name string
		run  func() ([]byte, error)
		want string
	}{
		{"foreign replacement", func() ([]byte, error) { return one.UniformReplace([]ElementReplacement{{Target: two.Elements()[1]}}) }, "foreign-target"},
		{"invalid authored name", func() ([]byte, error) {
			return one.UniformInsert([]ChildInsertion{{Parent: one.Root(), Children: []NewElement{{Name: xml.Name{Local: "bad:name"}}}}})
		}, "invalid-name-or-value"},
		{"unbound authored attribute namespace", func() ([]byte, error) {
			return one.UniformEdit(nil, []AttributeEdit{{Target: one.Elements()[1], Name: xml.Name{Space: "urn:unbound", Local: "a"}, Value: "v"}})
		}, "invalid-name-or-value"},
		{"invalid authored value", func() ([]byte, error) {
			return one.UniformInsert([]ChildInsertion{{Parent: one.Root(), Children: []NewElement{{Name: xml.Name{Local: "valid"}, Text: "\x01"}}}})
		}, "unsafe-XML"},
		{"duplicate empty insertion", func() ([]byte, error) {
			parent := one.Elements()[1]
			return one.UniformInsert([]ChildInsertion{{Parent: parent}, {Parent: parent}})
		}, "overlap-or-duplicate"},
		{"ancestor empty insertion", func() ([]byte, error) {
			return one.UniformInsert([]ChildInsertion{{Parent: one.Root()}, {Parent: one.Elements()[1]}})
		}, "overlap-or-duplicate"},
		{"foreign empty insertion", func() ([]byte, error) {
			return one.UniformInsert([]ChildInsertion{{Parent: two.Elements()[1]}})
		}, "foreign-target"},
		{"duplicate replacement", func() ([]byte, error) {
			e := one.Elements()[1]
			return one.UniformReplace([]ElementReplacement{{Target: e}, {Target: e}})
		}, "overlap-or-duplicate"},
		{"nested replacement", func() ([]byte, error) {
			return one.UniformReplace([]ElementReplacement{{Target: one.Root()}, {Target: one.Elements()[1]}})
		}, "root"},
	} {
		t.Run(test.name, func(t *testing.T) {
			out, e := test.run()
			if out != nil || UniformCategory(e) != test.want || !bytes.Equal(one.Source(), source) {
				t.Fatalf("out %q category %s err %v", out, UniformCategory(e), e)
			}
		})
	}
	if _, err := ParseUniform([]byte("<r>\x01</r>")); UniformCategory(err) != "unsafe-XML" {
		t.Errorf("invalid source character: %v", err)
	}
}
