package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func xmlNormalizationSteps(sc *godog.ScenarioContext) {
	var d *losslessxml.Document
	var target losslessxml.Element
	var source, output []byte
	var value string
	attr := func(e losslessxml.Element) (string, error) {
		for _, a := range e.Attributes() {
			if a.Name == (xml.Name{Local: "value"}) {
				return a.Value, nil
			}
		}
		return "", fmt.Errorf("value attribute missing")
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		d = nil
		source = nil
		output = nil
		value = ""
		return ctx, nil
	})
	sc.Step(`^XML text containing mixed CRLF and CR line endings$`, func() error {
		source = []byte("<r>\r\n<!--keep-->\r<t>A\r\nB\rC&#13;D</t>\r\n</r>")
		var err error
		d, err = losslessxml.Parse(source)
		if err == nil {
			target = d.Elements()[1]
		}
		return err
	})
	sc.Step(`^I inspect its decoded character data$`, func() error {
		var leaf bool
		value, leaf = target.Text()
		if !leaf {
			return fmt.Errorf("not a leaf")
		}
		return nil
	})
	sc.Step(`^literal line endings are LF while numeric CR remains CR$`, func() error {
		if value != "A\nB\nC\rD" {
			return fmt.Errorf("decoded %q", value)
		}
		return nil
	})
	sc.Step(`^a decoded-text no-op preserves the original bytes$`, func() error {
		out, err := d.Edit([]losslessxml.TextEdit{{Target: target, Text: value}}, nil)
		if err != nil {
			return err
		}
		if !bytes.Equal(out, source) {
			return fmt.Errorf("no-op normalized source bytes")
		}
		return nil
	})
	sc.Step(`^a changed text leaf preserves surrounding original line endings$`, func() error {
		out, err := d.Edit([]losslessxml.TextEdit{{Target: target, Text: "Changed"}}, nil)
		if err != nil {
			return err
		}
		want := bytes.Replace(source, []byte("A\r\nB\rC&#13;D"), []byte("Changed"), 1)
		if !bytes.Equal(out, want) {
			return fmt.Errorf("offset drift: %q", out)
		}
		return nil
	})
	sc.Step(`^an XML attribute with literal whitespace and numeric whitespace references$`, func() error {
		source = []byte("<r>\r\n<t keep = 'opaque' value = \"A\tB\nC\rD\r\nE&#9;F&#10;G&#13;H\">text</t>\r</r>")
		var err error
		d, err = losslessxml.Parse(source)
		if err == nil {
			target = d.Elements()[1]
		}
		return err
	})
	sc.Step(`^I inspect its decoded attribute value$`, func() error { var err error; value, err = attr(target); return err })
	sc.Step(`^literal attribute whitespace is spaces and numeric references retain their characters$`, func() error {
		if value != "A B C D E\tF\nG\rH" {
			return fmt.Errorf("decoded attribute %q", value)
		}
		return nil
	})
	sc.Step(`^a decoded-attribute no-op preserves its original lexical bytes$`, func() error {
		out, err := d.Edit(nil, []losslessxml.AttributeEdit{{Target: target, Name: xml.Name{Local: "value"}, Value: "A B C D E\tF\nG\rH"}})
		if err != nil {
			return err
		}
		if !bytes.Equal(out, source) {
			return fmt.Errorf("attribute no-op changed source")
		}
		return nil
	})
	sc.Step(`^I replace the attribute with tab CR and LF characters$`, func() error {
		var err error
		value = "x\ty\rz\nq"
		output, err = d.Edit(nil, []losslessxml.AttributeEdit{{Target: target, Name: xml.Name{Local: "value"}, Value: value}})
		return err
	})
	sc.Step(`^the changed attribute decodes to the requested characters$`, func() error {
		parsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		got, err := attr(parsed.Elements()[1])
		if err != nil {
			return err
		}
		if got != value {
			return fmt.Errorf("roundtrip %q != %q", got, value)
		}
		return nil
	})
	sc.Step(`^unrelated attributes and start-tag spacing remain byte-identical$`, func() error {
		prefix := []byte("<r>\r\n<t keep = 'opaque' value = \"")
		suffix := []byte("\">text</t>\r</r>")
		if !bytes.HasPrefix(output, prefix) || !bytes.HasSuffix(output, suffix) {
			return fmt.Errorf("unrelated bytes changed")
		}
		return nil
	})
}
