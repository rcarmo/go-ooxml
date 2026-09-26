package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func graphEditSteps(sc *godog.ScenarioContext) {
	var p *packaging.Preserved
	var plan *packaging.GraphPlan
	var source, output []byte
	var failure error
	const ct = `<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"><ct:Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><!--types--><ct:Default Extension='bin' ContentType='application/octet-stream'/></ct:Types>`
	const rels = `<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"><!--keep--><r:Relationship Id = 'first' Type='urn:image' Target = 'media/shared.bin'/><r:Relationship Id='second' Type='urn:image' Target='media/shared.bin'/><r:Relationship Id='external' Type='urn:link' Target='https://example.invalid/' TargetMode='External'/></r:Relationships>`
	serialize := func() ([]byte, error) { var b bytes.Buffer; err := p.WriteTo(&b); return b.Bytes(), err }
	change := func() packaging.GraphMutation {
		return packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: "media/private.bin", ContentType: "application/octet-stream", Data: []byte("new payload")}}, Retargets: []packaging.RelationshipRetarget{{Source: "", ID: "first", TargetPart: "media/private.bin"}}}
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		p = nil
		plan = nil
		failure = nil
		source = nil
		output = nil
		return ctx, nil
	})
	sc.Step(`^a retained package with two edges sharing a payload$`, func() error {
		var err error
		var b bytes.Buffer
		z := zip.NewWriter(&b)
		for _, m := range []struct{ name, text string }{{"[Content_Types].xml", ct}, {"_rels/.rels", rels}, {"media/shared.bin", "old payload"}, {"opaque.bin", "sentinel"}} {
			w, e := z.Create(m.name)
			if e != nil {
				return e
			}
			if _, e = w.Write([]byte(m.text)); e != nil {
				return e
			}
		}
		if err = z.Close(); err != nil {
			return err
		}
		source = b.Bytes()
		p, err = packaging.OpenPreserved(source, packaging.Limits{})
		return err
	})
	sc.Step(`^I plan a new payload for the first edge$`, func() error { var err error; plan, err = p.PlanGraphMutation(change()); return err })
	sc.Step(`^planning leaves the original archive byte-identical$`, func() error {
		out, err := serialize()
		if err != nil {
			return err
		}
		if !bytes.Equal(source, out) {
			return fmt.Errorf("planning mutated package")
		}
		return nil
	})
	sc.Step(`^I apply and deliver the graph plan$`, func() error {
		if err := p.ApplyGraphPlan(plan); err != nil {
			return err
		}
		dir, err := os.MkdirTemp("", "graph-plan-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.zip")
		r, err := p.SaveAs(path)
		if err != nil {
			return err
		}
		if r.Schema != 2 || len(r.Changes) != 3 {
			return fmt.Errorf("wrong addition receipt %+v", r)
		}
		output, err = os.ReadFile(path)
		return err
	})
	sc.Step(`^only the selected edge points to the new registered payload$`, func() error {
		q, err := packaging.OpenPreserved(output, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		found := map[string]string{}
		for _, e := range g.Edges {
			found[e.ID] = e.ResolvedPart
		}
		if found["first"] != "media/private.bin" || found["second"] != "media/shared.bin" {
			return fmt.Errorf("edge mismatch %+v", found)
		}
		parts, _ := zipPayloads(output)
		if string(parts["media/private.bin"]) != "new payload" {
			return fmt.Errorf("new data absent")
		}
		return nil
	})
	sc.Step(`^unrelated registry bytes and original payloads are preserved$`, func() error {
		before, _ := zipPayloads(source)
		after, err := zipPayloads(output)
		if err != nil {
			return err
		}
		want := bytes.Replace(before["_rels/.rels"], []byte(`Target = 'media/shared.bin'`), []byte(`Target = 'media/private.bin'`), 1)
		if !bytes.Equal(want, after["_rels/.rels"]) {
			return fmt.Errorf("registry reserialized")
		}
		for _, name := range []string{"media/shared.bin", "opaque.bin"} {
			if !bytes.Equal(before[name], after[name]) {
				return fmt.Errorf("original payload changed")
			}
		}
		wantCT := bytes.Replace(before["[Content_Types].xml"], []byte(`</ct:Types>`), []byte(`<ct:Override PartName="/media/private.bin" ContentType="application/octet-stream"/></ct:Types>`), 1)
		if !bytes.Equal(wantCT, after["[Content_Types].xml"]) {
			return fmt.Errorf("types reserialized")
		}
		return nil
	})
	sc.Step(`^I plan a graph change with "([^"]+)"$`, func(condition string) error {
		c := change()
		switch condition {
		case "case colliding part":
			c.Additions[0].Name = "MEDIA/SHARED.bin"
		case "absent target":
			c.Retargets[0].TargetPart = "missing.bin"
		case "external edge":
			c.Retargets[0].ID = "external"
		case "duplicate edge edit":
			c.Retargets = append(c.Retargets, c.Retargets[0])
		case "reserved registry addition":
			c.Additions[0].Name = "new/_rels/a.xml.rels"
		case "malformed XML addition":
			c.Additions[0].Name = "bad.xml"
			c.Additions[0].Data = []byte(`<bad>`)
			c.Retargets[0].TargetPart = "bad.xml"
		default:
			return fmt.Errorf("unknown condition")
		}
		_, failure = p.PlanGraphMutation(c)
		return nil
	})
	sc.Step(`^graph planning refuses without adding parts or changing bytes$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("missing refusal %v", failure)
		}
		out, err := serialize()
		if err != nil {
			return err
		}
		if !bytes.Equal(source, out) {
			return fmt.Errorf("refusal changed archive")
		}
		return nil
	})
	sc.Step(`^I change an existing payload before applying the graph plan$`, func() error {
		_, hash, err := p.Part("opaque.bin")
		if err != nil {
			return err
		}
		return p.Replace([]packaging.Replacement{{Part: "opaque.bin", ExpectedSHA256: hash, Data: []byte("intervening")}})
	})
	sc.Step(`^the stale graph plan refuses and preserves the intervening edit$`, func() error {
		before, err := serialize()
		if err != nil {
			return err
		}
		err = p.ApplyGraphPlan(plan)
		var r *packaging.Refusal
		if !errors.As(err, &r) || r.Kind != "stale_target" {
			return fmt.Errorf("expected stale refusal: %v", err)
		}
		after, err := serialize()
		if err != nil {
			return err
		}
		if !bytes.Equal(before, after) {
			return fmt.Errorf("stale apply mutated")
		}
		return nil
	})
}
