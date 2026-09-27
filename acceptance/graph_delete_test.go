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

func graphDeleteSteps(sc *godog.ScenarioContext) {
	var p *packaging.Preserved
	var source, out []byte
	var plan *packaging.GraphPlan
	var failure error
	const ct = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><!--keep types--><Override PartName='/leaf.bin' ContentType='application/octet-stream'/></Types>`
	const a = `<Relationship Id='a' Type='urn:test' Target='leaf.bin'/>`
	const b = `<Relationship Id='b' Type='urn:test' Target='leaf.bin'/>`
	const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><!--keep edges-->` + a + b + `</Relationships>`
	serialize := func() ([]byte, error) { var b bytes.Buffer; err := p.WriteTo(&b); return b.Bytes(), err }
	request := func() packaging.GraphMutation {
		_, leaf, _ := p.Part("leaf.bin")
		_, owner, _ := p.Part("owner.xml")
		return packaging.GraphMutation{Removals: []packaging.RelationshipRemoval{{Source: "", ID: "a"}, {Source: "", ID: "b"}}, Deletions: []packaging.PartDeletion{{Name: "leaf.bin", ExpectedSHA256: leaf}}, Replacements: []packaging.Replacement{{Part: "owner.xml", ExpectedSHA256: owner, Data: []byte(`<owner>new</owner>`)}}}
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		p = nil
		plan = nil
		failure = nil
		source = nil
		out = nil
		return ctx, nil
	})
	sc.Step(`^a package with two relationships to a removable leaf$`, func() error {
		var data bytes.Buffer
		z := zip.NewWriter(&data)
		for _, m := range []struct{ name, text string }{{"[Content_Types].xml", ct}, {"_rels/.rels", rels}, {"leaf.bin", "leaf"}, {"owner.xml", "<owner>old</owner>"}, {"keep.xml", "<keep/>"}} {
			w, err := z.Create(m.name)
			if err != nil {
				return err
			}
			if _, err = w.Write([]byte(m.text)); err != nil {
				return err
			}
		}
		if err := z.Close(); err != nil {
			return err
		}
		source = data.Bytes()
		var err error
		p, err = packaging.OpenPreserved(source, packaging.Limits{})
		return err
	})
	sc.Step(`^I plan removal of both edges and the leaf with an owner edit$`, func() error { var err error; plan, err = p.PlanGraphMutation(request()); return err })
	sc.Step(`^the deletion preview preserves all package bytes$`, func() error {
		got, err := serialize()
		if err != nil {
			return err
		}
		if !bytes.Equal(source, got) {
			return fmt.Errorf("preview mutated")
		}
		return nil
	})
	sc.Step(`^I apply and deliver the deletion plan$`, func() error {
		if err := p.ApplyGraphPlan(plan); err != nil {
			return err
		}
		dir, err := os.MkdirTemp("", "delete-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "out.zip")
		receipt, err := p.SaveAs(path)
		if err != nil {
			return err
		}
		found := false
		for _, c := range receipt.Changes {
			if c.Part == "leaf.bin" && c.Operation == "delete" && c.BeforeSHA256 != "" && c.AfterSHA256 == "" {
				found = true
			}
		}
		if receipt.Schema != 2 || !found {
			return fmt.Errorf("missing delete receipt %+v", receipt)
		}
		out, err = os.ReadFile(path)
		return err
	})
	sc.Step(`^the leaf and its content-type override and edges are absent$`, func() error {
		q, err := packaging.OpenPreserved(out, packaging.Limits{})
		if err != nil {
			return err
		}
		if _, _, err = q.Part("leaf.bin"); err == nil {
			return fmt.Errorf("leaf remains")
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		if len(g.Edges) != 0 {
			return fmt.Errorf("edges remain")
		}
		parts, _ := zipPayloads(out)
		if bytes.Contains(parts["[Content_Types].xml"], []byte("leaf.bin")) {
			return fmt.Errorf("override remains")
		}
		return nil
	})
	sc.Step(`^the owner edit and unrelated registry bytes are preserved$`, func() error {
		parts, err := zipPayloads(out)
		if err != nil {
			return err
		}
		want := bytes.Replace([]byte(ct), []byte(`<Override PartName='/leaf.bin' ContentType='application/octet-stream'/>`), nil, 1)
		wantRel := bytes.ReplaceAll(bytes.ReplaceAll([]byte(rels), []byte(a), nil), []byte(b), nil)
		if string(parts["owner.xml"]) != "<owner>new</owner>" || string(parts["keep.xml"]) != "<keep/>" || !bytes.Equal(parts["[Content_Types].xml"], want) || !bytes.Equal(parts["_rels/.rels"], wantRel) {
			return fmt.Errorf("owner/registry payload mismatch")
		}
		return nil
	})
	sc.Step(`^I attempt a deletion plan with "([^"]+)"$`, func(condition string) error {
		c := request()
		switch condition {
		case "remaining inbound edge":
			c.Removals = c.Removals[:1]
		case "stale leaf fingerprint":
			c.Deletions[0].ExpectedSHA256 = "stale"
		case "stale owner replacement":
			c.Replacements[0].ExpectedSHA256 = "stale"
		case "duplicate edge removal":
			c.Removals = append(c.Removals, c.Removals[0])
		case "delete and replace conflict":
			c.Replacements = append(c.Replacements, packaging.Replacement{Part: "leaf.bin", ExpectedSHA256: c.Deletions[0].ExpectedSHA256, Data: []byte("other")})
		case "missing leaf":
			c.Deletions[0].Name = "missing.bin"
		default:
			return fmt.Errorf("unknown condition")
		}
		_, failure = p.PlanGraphMutation(c)
		return nil
	})
	sc.Step(`^deletion planning refuses with unchanged package bytes$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) {
			return fmt.Errorf("missing refusal %v", failure)
		}
		got, err := serialize()
		if err != nil {
			return err
		}
		if !bytes.Equal(source, got) {
			return fmt.Errorf("failed plan changed bytes")
		}
		return nil
	})
}
