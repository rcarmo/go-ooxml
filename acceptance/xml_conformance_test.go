package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func xmlConformanceSteps(sc *godog.ScenarioContext) {
	var source, original, output []byte
	var failure error
	var d *losslessxml.Document
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		original = nil
		output = nil
		failure = nil
		d = nil
		return ctx, nil
	})
	sc.Step(`^XML with invalid qualified-name case "([^"]+)"$`, func(which string) error {
		cases := map[string]string{"numeric local name": `<p:1 xmlns:p="u"/>`, "numeric attribute local name": `<r xmlns:p="u" p:1="v"/>`, "numeric namespace prefix": `<r xmlns:1="u"><1:x/></r>`, "non XML whitespace outside root": "\u00a0<r/>"}
		s, ok := cases[which]
		if !ok {
			return fmt.Errorf("unknown case")
		}
		source = []byte(s)
		original = bytes.Clone(source)
		return nil
	})
	sc.Step(`^I parse it as an editable XML snapshot$`, func() error { _, failure = losslessxml.Parse(source); return nil })
	sc.Step(`^invalid XML structure is refused without changing the input$`, func() error {
		if failure == nil {
			return fmt.Errorf("invalid XML accepted")
		}
		if !bytes.Equal(source, original) {
			return fmt.Errorf("parser changed input")
		}
		return nil
	})
	sc.Step(`^a parent with an occupied generated prefix and default namespace$`, func() error {
		source = []byte(`<r xmlns="urn:old" xmlns:n1="urn:occupied"><n1:existing/></r>`)
		var err error
		d, err = losslessxml.Parse(source)
		return err
	})
	sc.Step(`^I append children with independent new namespace requirements$`, func() error {
		var err error
		output, err = d.InsertChildren([]losslessxml.ChildInsertion{{Parent: d.Elements()[0], Children: []losslessxml.NewElement{{Name: xml.Name{Space: "urn:new", Local: "first"}, Attributes: []xml.Attr{{Name: xml.Name{Space: "urn:attr", Local: "flag"}, Value: "v"}}}, {Name: xml.Name{Local: "plain"}}, {Name: xml.Name{Space: "urn:old", Local: "last"}}}}})
		return err
	})
	sc.Step(`^all new and existing sibling names resolve to their intended namespaces$`, func() error {
		parsed, err := losslessxml.Parse(output)
		if err != nil {
			return err
		}
		want := []xml.Name{{Space: "urn:old", Local: "r"}, {Space: "urn:occupied", Local: "existing"}, {Space: "urn:new", Local: "first"}, {Local: "plain"}, {Space: "urn:old", Local: "last"}}
		elements := parsed.Elements()
		if len(elements) != len(want) {
			return fmt.Errorf("wrong node inventory")
		}
		for i, e := range elements {
			if e.Name() != want[i] {
				return fmt.Errorf("node%d %+v expected%+v", i, e.Name(), want[i])
			}
		}
		if !bytes.Contains(output, []byte(`<n1:existing/>`)) {
			return fmt.Errorf("existing QName changed")
		}
		return nil
	})
}
