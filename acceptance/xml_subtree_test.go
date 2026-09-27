package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func subtreeSteps(sc *godog.ScenarioContext) {
	var d *losslessxml.Document
	var source, out []byte
	var failure error
	nodes := func() []losslessxml.NewElement {
		return []losslessxml.NewElement{{Name: xml.Name{Space: "outer", Local: "new"}, Text: "A&B"}, {Name: xml.Name{Local: "plain"}}}
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		d = nil
		source = nil
		out = nil
		failure = nil
		return ctx, nil
	})
	sc.Step(`^an XML subtree that shadows its parent's prefix$`, func() error {
		source = []byte(`<r xmlns="default" xmlns:p="outer">before<!--keep--><p:old xmlns:p="inner"/> after</r>`)
		var err error
		d, err = losslessxml.Parse(source)
		return err
	})
	sc.Step(`^I replace it with structured sibling nodes$`, func() error {
		var err error
		out, err = d.ReplaceElements([]losslessxml.ElementReplacement{{Target: d.Elements()[1], Nodes: nodes()}})
		return err
	})
	sc.Step(`^inserted names use the parent scope and surrounding bytes stay exact$`, func() error {
		want := bytes.Replace(source, []byte(`<p:old xmlns:p="inner"/>`), []byte(`<p:new>A&amp;B</p:new><plain xmlns=""/>`), 1)
		if !bytes.Equal(out, want) {
			return fmt.Errorf("unexpected subtree output %s", out)
		}
		p, err := losslessxml.Parse(out)
		if err != nil {
			return err
		}
		if p.Elements()[1].Name() != (xml.Name{Space: "outer", Local: "new"}) || p.Elements()[2].Name() != (xml.Name{Local: "plain"}) {
			return fmt.Errorf("namespace changed")
		}
		return nil
	})
	sc.Step(`^one of the structured sibling replacements contains invalid XML text$`, func() error {
		n := nodes()
		n[1].Text = "bad\x00"
		out, failure = d.ReplaceElements([]losslessxml.ElementReplacement{{Target: d.Elements()[1], Nodes: n}})
		return nil
	})
	sc.Step(`^subtree replacement refuses without publishing partial output$`, func() error {
		if failure == nil || out != nil {
			return fmt.Errorf("partial/successful output")
		}
		got, err := d.Edit(nil, nil)
		if err != nil {
			return err
		}
		if !bytes.Equal(source, got) {
			return fmt.Errorf("snapshot changed")
		}
		return nil
	})
}
