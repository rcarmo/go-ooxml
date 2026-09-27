package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func attributeSteps(sc *godog.ScenarioContext) {
	const source = `<r xmlns:q="urn:q"><q:t a = 'keep' q:kind="q:Type">old</q:t></r>`
	var d *losslessxml.Document
	var target losslessxml.Element
	var output []byte
	var failure error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		d = nil
		output = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^XML with a text element and unrelated attributes$`, func() error {
		var err error
		d, err = losslessxml.Parse([]byte(source))
		if err != nil {
			return err
		}
		target = d.Elements()[1]
		return nil
	})
	sc.Step(`^I add XML space preservation and replace its text$`, func() error {
		var err error
		output, err = d.Edit([]losslessxml.TextEdit{{Target: target, Text: " new "}}, []losslessxml.AttributeEdit{{Target: target, Name: xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}, Value: "preserve"}})
		return err
	})
	sc.Step(`^the original start-tag attributes retain their bytes$`, func() error {
		if !bytes.Contains(output, []byte(`a = 'keep' q:kind="q:Type"`)) {
			return fmt.Errorf("attributes rewritten")
		}
		return nil
	})
	sc.Step(`^only the requested attribute and text are changed$`, func() error {
		want := `<r xmlns:q="urn:q"><q:t a = 'keep' q:kind="q:Type" xml:space="preserve"> new </q:t></r>`
		if string(output) != want {
			return fmt.Errorf("unexpected %s", output)
		}
		return nil
	})
	sc.Step(`^an attribute edit attempts to change a namespace declaration$`, func() error {
		_, failure = d.Edit(nil, []losslessxml.AttributeEdit{{Target: target, Name: xml.Name{Local: "xmlns"}, Value: "urn:other"}})
		return nil
	})
	sc.Step(`^the attribute batch refuses and the original snapshot remains reusable$`, func() error {
		if failure == nil {
			return fmt.Errorf("namespace change accepted")
		}
		b, err := d.Edit(nil, nil)
		if err != nil {
			return err
		}
		if string(b) != source {
			return fmt.Errorf("snapshot changed")
		}
		_, err = d.Edit(nil, []losslessxml.AttributeEdit{{Target: target, Name: xml.Name{Local: "a"}, Value: "new"}})
		return err
	})
}
