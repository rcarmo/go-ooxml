package acceptance

import (
	"bytes"
	"context"
	"fmt"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func graphTargetFormSteps(sc *godog.ScenarioContext) {
	var p *packaging.Preserved
	var source, output []byte
	const target = "xl/media/new & 雪.bin"
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		p = nil
		source = nil
		output = nil
		return ctx, nil
	})
	sc.Step(`^a graph edge using "([^"]+)" target form$`, func(form string) error {
		q := packaging.New()
		_, _ = q.AddPart("xl/drawings/d.xml", packaging.ContentTypeDrawing, []byte(`<r/>`))
		_, _ = q.AddPart("xl/media/old.bin", "application/octet-stream", []byte("old"))
		q.AddRelationship("", "xl/drawings/d.xml", packaging.RelTypeOfficeDocument)
		path := "../media/old.bin"
		if form == "absolute" {
			path = "/xl/media/old.bin"
		} else if form != "relative" {
			return fmt.Errorf("unknown form")
		}
		q.AddRelationship("xl/drawings/d.xml", path, packaging.RelTypeImage)
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		var err error
		p, err = packaging.OpenPreserved(source, packaging.Limits{})
		return err
	})
	sc.Step(`^I retarget it to media containing escaped characters$`, func() error {
		plan, err := p.PlanGraphMutation(packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: target, ContentType: "application/octet-stream", Data: []byte("new")}}, Retargets: []packaging.RelationshipRetarget{{Source: "xl/drawings/d.xml", ID: "rId1", TargetPart: target}}})
		if err != nil {
			return err
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			return err
		}
		var b bytes.Buffer
		if err = p.WriteTo(&b); err != nil {
			return err
		}
		output = b.Bytes()
		return nil
	})
	sc.Step(`^its new target has "([^"]+)" form and resolves to the exact added member$`, func(form string) error {
		q, err := packaging.OpenPreserved(output, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		for _, e := range g.Edges {
			if e.Source != "xl/drawings/d.xml" {
				continue
			}
			if e.ResolvedPart != target || strings.HasPrefix(e.Target, "/") != (form == "absolute") || !strings.Contains(e.Target, "%20") {
				return fmt.Errorf("wrong target form %+v", e)
			}
			return nil
		}
		return fmt.Errorf("edge absent")
	})
	sc.Step(`^the old media and all unrelated payloads remain unchanged$`, func() error {
		a, _ := zipPayloads(source)
		b, err := zipPayloads(output)
		if err != nil {
			return err
		}
		for name, data := range a {
			if name == "[Content_Types].xml" || name == "xl/drawings/_rels/d.xml.rels" {
				continue
			}
			if !bytes.Equal(data, b[name]) {
				return fmt.Errorf("part changed %s", name)
			}
		}
		return nil
	})
}
