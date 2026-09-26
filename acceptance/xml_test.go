package acceptance

import (
	"bytes"
	"context"
	"fmt"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

const xmlSample = `<?xml version="1.0"?><r xmlns:w='urn:word' xmlns:x="urn:opaque" x:type='x:Kind'><!--keep--><w:t xml:space="preserve">old &amp; value</w:t><x:unknown z = '1'/></r>`

func xmlSteps(sc *godog.ScenarioContext) {
	var source, original, output []byte
	var d *losslessxml.Document
	var parseErr, errorBatch error
	target := func(doc *losslessxml.Document) (losslessxml.Element, error) {
		for _, e := range doc.Elements() {
			if e.Name().Space == "urn:word" && e.Name().Local == "t" {
				return e, nil
			}
		}
		return losslessxml.Element{}, fmt.Errorf("missing expanded-name leaf")
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		original = nil
		output = nil
		d = nil
		parseErr = nil
		errorBatch = nil
		return ctx, nil
	})
	sc.Step(`^XML with prefixed text and an opaque extension$`, func() error {
		source = []byte(xmlSample)
		original = bytes.Clone(source)
		var err error
		d, err = losslessxml.Parse(source)
		return err
	})
	sc.Step(`^I replace the selected text leaf with "([^"]+)"$`, func(text string) error {
		e, err := target(d)
		if err != nil {
			return err
		}
		output, err = d.ReplaceText([]losslessxml.TextEdit{{Target: e, Text: text}})
		return err
	})
	sc.Step(`^only the leaf character-data bytes change$`, func() error {
		expected := strings.Replace(xmlSample, "old &amp; value", "A&amp;B &lt;new&gt;", 1)
		if string(output) != expected {
			return fmt.Errorf("unexpected XML bytes %s", output)
		}
		if !bytes.Equal(source, original) {
			return fmt.Errorf("source changed")
		}
		return nil
	})
	sc.Step(`^the replacement decodes to "([^"]+)"$`, func(text string) error {
		parsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		e, err := target(parsed)
		if err != nil {
			return err
		}
		got, ok := e.Text()
		if !ok || got != text {
			return fmt.Errorf("text %q leaf=%v", got, ok)
		}
		return nil
	})
	sc.Step(`^XML with "([^"]+)"$`, func(defect string) error {
		cases := map[string]string{"undeclared element prefix": `<r><w:t>bad</w:t></r>`, "undeclared attribute prefix": `<r x:a="bad"/>`, "duplicate expanded attribute": `<r xmlns:a="urn:x" xmlns:b="urn:x" a:n="1" b:n="2"/>`, "invalid xml binding": `<r xmlns:xml="urn:wrong"/>`, "mismatched end prefix": `<a:r xmlns:a="urn:x" xmlns:b="urn:x"></b:r>`}
		s, ok := cases[defect]
		if !ok {
			return fmt.Errorf("unknown defect")
		}
		source = []byte(s)
		original = bytes.Clone(source)
		return nil
	})
	sc.Step(`^I parse the XML for lossless editing$`, func() error { d, parseErr = losslessxml.Parse(source); return nil })
	sc.Step(`^parsing refuses without changing the source$`, func() error {
		if parseErr == nil {
			return fmt.Errorf("invalid namespaces accepted")
		}
		if !bytes.Equal(source, original) {
			return fmt.Errorf("source changed")
		}
		return nil
	})
	sc.Step(`^an edit batch contains a target from another parsed document$`, func() error {
		e, err := target(d)
		if err != nil {
			return err
		}
		other, err := losslessxml.Parse(source)
		if err != nil {
			return err
		}
		foreign, err := target(other)
		if err != nil {
			return err
		}
		_, errorBatch = d.ReplaceText([]losslessxml.TextEdit{{e, "new"}, {foreign, "bad"}})
		return nil
	})
	sc.Step(`^the batch refuses and a valid target remains reusable$`, func() error {
		if errorBatch == nil {
			return fmt.Errorf("foreign target accepted")
		}
		e, err := target(d)
		if err != nil {
			return err
		}
		b, err := d.ReplaceText([]losslessxml.TextEdit{{e, "new"}})
		if err != nil {
			return err
		}
		if string(b) != strings.Replace(xmlSample, "old &amp; value", "new", 1) {
			return fmt.Errorf("target corrupted")
		}
		return nil
	})
}
