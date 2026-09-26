package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"fmt"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func graphSteps(sc *godog.ScenarioContext) {
	var source []byte
	var p *packaging.Preserved
	var graph packaging.Graph
	var failure error
	setup := func(defect string) error {
		q := packaging.New()
		for _, n := range []string{"word/document.xml", "word/header.xml"} {
			if _, err := q.AddPart(n, packaging.ContentTypeXML, []byte(`<r/>`)); err != nil {
				return err
			}
			q.AddRelationship(n, "media/shared.png", packaging.RelTypeImage)
		}
		if _, err := q.AddPart("word/media/shared.png", packaging.ContentTypePNG, []byte("image")); err != nil {
			return err
		}
		q.AddRelationship("", "word/document.xml", packaging.RelTypeOfficeDocument)
		q.AddRelationshipWithTargetMode("word/document.xml", "https://example.invalid/no-fetch", packaging.RelTypeHyperlink, packaging.TargetModeExternal)
		switch defect {
		case "missing target":
			_ = q.DeletePart("word/media/shared.png")
		case "duplicate relationship ID":
			rels := q.GetRelationships("word/document.xml")
			rels.Relationships = append(rels.Relationships, rels.Relationships[0])
		case "escaping target":
			q.GetRelationships("word/document.xml").Relationships[0].Target = "../../outside.xml"
		case "duplicate content-type default":
			q.ContentTypes().Defaults = append(q.ContentTypes().Defaults, q.ContentTypes().Defaults[0])
		}
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			return err
		}
		source = b.Bytes()
		if defect == "unknown registry extension" {
			r, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
			if err != nil {
				return err
			}
			var modified bytes.Buffer
			z := zip.NewWriter(&modified)
			parts, err := zipPayloads(source)
			if err != nil {
				return err
			}
			for _, f := range r.File {
				data := parts[f.Name]
				if f.Name == "word/_rels/document.xml.rels" {
					data = []byte(strings.Replace(string(data), "</Relationships>", `<extra xmlns="urn:unknown"/></Relationships>`, 1))
				}
				w, err := z.Create(f.Name)
				if err != nil {
					return err
				}
				if _, err = w.Write(data); err != nil {
					return err
				}
			}
			if err = z.Close(); err != nil {
				return err
			}
			source = modified.Bytes()
		}
		var err error
		p, err = packaging.OpenPreserved(source, packaging.Limits{})
		return err
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source = nil
		p = nil
		graph = packaging.Graph{}
		failure = nil
		return ctx, nil
	})
	sc.Step(`^a package with two users of one image and an external hyperlink$`, func() error { return setup("") })
	sc.Step(`^a relationship package with "([^"]+)"$`, setup)
	sc.Step(`^I inspect the retained package relationship graph$`, func() error { graph, failure = p.Graph(); return nil })
	sc.Step(`^the image has two incoming relationships$`, func() error {
		if failure != nil {
			return failure
		}
		for _, part := range graph.Parts {
			if part.Name == "word/media/shared.png" && part.Inbound == 2 {
				return nil
			}
		}
		return fmt.Errorf("shared ownership missing")
	})
	sc.Step(`^the hyperlink remains an unresolved external edge$`, func() error {
		for _, e := range graph.Edges {
			if e.Target == "https://example.invalid/no-fetch" && e.External && e.ResolvedPart == "" {
				return nil
			}
		}
		return fmt.Errorf("external edge changed")
	})
	sc.Step(`^graph inspection leaves the archive byte-identical$`, func() error {
		var b bytes.Buffer
		if err := p.WriteTo(&b); err != nil {
			return err
		}
		if !bytes.Equal(b.Bytes(), source) {
			return fmt.Errorf("inspection mutated source")
		}
		return nil
	})
	sc.Step(`^graph inspection returns a relationship-policy refusal$`, func() error {
		var r *packaging.Refusal
		if !errors.As(failure, &r) || r.Kind != "relationship_policy" {
			return fmt.Errorf("expected relationship_policy, got %v", failure)
		}
		return nil
	})
}
