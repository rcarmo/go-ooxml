package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func xmlInsertSteps(sc *godog.ScenarioContext) {
	var source, output []byte
	var d *losslessxml.Document
	var parent losslessxml.Element
	var failure error
	parse := func(s string) error {
		source = []byte(s)
		var err error
		d, err = losslessxml.Parse(source)
		if err == nil {
			parent = d.Elements()[0]
		}
		return err
	}
	child := func() losslessxml.NewElement {
		return losslessxml.NewElement{Name: xml.Name{Space: "urn:p", Local: "child"}, Attributes: []xml.Attr{{Name: xml.Name{Space: "urn:a", Local: "kind"}, Value: "x&y"}}, Text: "A<B"}
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		output = nil
		d = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a prefixed XML parent with opaque surrounding markup$`, func() error {
		return parse(`<p:root xmlns:p='urn:p' xmlns:a="urn:a"><!--opaque--><a:old key = 'keep'/></p:root>`)
	})
	sc.Step(`^an XML parent in a default namespace$`, func() error { return parse(`<root xmlns="urn:default"><existing/></root>`) })
	sc.Step(`^a self-closing prefixed parent with unusual attribute spacing$`, func() error { return parse(`<p:root xmlns:p='urn:p' xmlns:a="urn:a" key = 'keep' />`) })
	sc.Step(`^I append a same-namespace child with a namespaced attribute$`, func() error {
		var err error
		output, err = d.InsertChildren([]losslessxml.ChildInsertion{{Parent: parent, Children: []losslessxml.NewElement{child()}}})
		return err
	})
	sc.Step(`^the new child and attribute have the requested expanded names$`, func() error {
		parsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		found := 0
		for _, e := range parsed.Elements() {
			if e.Name() == (xml.Name{Space: "urn:p", Local: "child"}) {
				found++
				text, leaf := e.Text()
				if !leaf || text != "A<B" {
					return fmt.Errorf("wrong child text")
				}
				attrs := e.Attributes()
				if len(attrs) != 1 || attrs[0].Name != (xml.Name{Space: "urn:a", Local: "kind"}) || attrs[0].Value != "x&y" {
					return fmt.Errorf("wrong attribute %+v", attrs)
				}
			}
		}
		if found != 1 {
			return fmt.Errorf("child identity missing")
		}
		return nil
	})
	sc.Step(`^every original byte outside the insertion remains identical$`, func() error {
		insert := []byte(`<p:child a:kind="x&amp;y">A&lt;B</p:child>`)
		want := bytes.Replace(source, []byte(`</p:root>`), append(insert, []byte(`</p:root>`)...), 1)
		if !bytes.Equal(output, want) {
			return fmt.Errorf("unexpected output %s", output)
		}
		return nil
	})
	sc.Step(`^I append an explicitly no-namespace child$`, func() error {
		var err error
		output, err = d.InsertChildren([]losslessxml.ChildInsertion{{Parent: parent, Children: []losslessxml.NewElement{{Name: xml.Name{Local: "plain"}, Text: "v"}}}})
		return err
	})
	sc.Step(`^the child has no namespace and the parent keeps its namespace$`, func() error {
		parsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		es := parsed.Elements()
		if len(es) != 3 || es[0].Name().Space != "urn:default" || es[2].Name() != (xml.Name{Local: "plain"}) {
			return fmt.Errorf("default namespace leaked")
		}
		return nil
	})
	sc.Step(`^the parent attributes retain their original lexical bytes$`, func() error {
		if !bytes.HasPrefix(output, []byte(`<p:root xmlns:p='urn:p' xmlns:a="urn:a" key = 'keep' >`)) {
			return fmt.Errorf("parent rewritten %s", output)
		}
		return nil
	})
	sc.Step(`^one insertion attempts to supply a namespace declaration$`, func() error {
		bad := child()
		bad.Attributes = append(bad.Attributes, xml.Attr{Name: xml.Name{Local: "xmlns"}, Value: "urn:bad"})
		_, failure = d.InsertChildren([]losslessxml.ChildInsertion{{Parent: parent, Children: []losslessxml.NewElement{child(), bad}}})
		return nil
	})
	sc.Step(`^the insertion batch refuses with a reusable unchanged snapshot$`, func() error {
		if failure == nil {
			return fmt.Errorf("namespace mutation accepted")
		}
		b, err := d.Edit(nil, nil)
		if err != nil {
			return err
		}
		if !bytes.Equal(source, b) {
			return fmt.Errorf("snapshot mutated")
		}
		_, err = d.InsertChildren([]losslessxml.ChildInsertion{{Parent: parent, Children: []losslessxml.NewElement{child()}}})
		return err
	})
}
