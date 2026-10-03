package packaging

import (
	"bytes"
	"testing"
)

func TestFreshRelationshipRegistryGraphPlan(t *testing.T) {
	p := graphEditPackage(t)
	a, b := "word/new/data.xml", "word/new/drawing.xml"
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="drawing" Type="urn:diagram" Target="drawing.xml"/><Relationship Id="external" Type="urn:link" Target="https://invalid.example/a?x=1&amp;y=2" TargetMode="External"/></Relationships>`
	reverse := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="back" Type="urn:diagram" Target="data.xml"/></Relationships>`
	plan, e := p.PlanGraphMutation(GraphMutation{Additions: []PartAddition{{Name: a, ContentType: "application/xml", Data: []byte(`<a/>`)}, {Name: b, ContentType: "application/xml", Data: []byte(`<b/>`)}}, Registries: []RelationshipRegistryAddition{{Owner: a, Data: []byte(rels)}, {Owner: b, Data: []byte(reverse)}}})
	if e != nil {
		t.Fatal(e)
	}
	if e = p.ApplyGraphPlan(plan); e != nil {
		t.Fatal(e)
	}
	graph, e := p.Graph()
	if e != nil {
		t.Fatal(e)
	}
	found := 0
	for _, edge := range graph.Edges {
		if edge.Source == a || edge.Source == b {
			found++
			if edge.External && edge.Target != "https://invalid.example/a?x=1&y=2" {
				t.Fatal("external spelling")
			}
		}
	}
	if found != 3 {
		t.Fatal("cycle registry closure")
	}
	raw, _, e := p.Part(RelationshipsPathForPart(a))
	if e != nil || string(raw) != rels {
		t.Fatal("registry lexical preservation", e)
	}
}
func TestFreshRegistryRefusalsAreAtomic(t *testing.T) {
	for _, c := range []struct{ name, owner, xml string }{{"existing", "word/document.xml", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`}, {"dangling", "word/new.xml", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="urn:x" Target="absent.xml"/></Relationships>`}, {"wrong-root", "word/new.xml", `<wrong/>`}, {"bad-owner", "../bad.xml", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`}} {
		t.Run(c.name, func(t *testing.T) {
			p := graphEditPackage(t)
			var initial bytes.Buffer
			if e := p.WriteTo(&initial); e != nil {
				t.Fatal(e)
			}
			source := initial.Bytes()
			_, e := p.PlanGraphMutation(GraphMutation{Additions: []PartAddition{{Name: "word/new.xml", ContentType: "application/xml", Data: []byte(`<new/>`)}}, Registries: []RelationshipRegistryAddition{{Owner: c.owner, Data: []byte(c.xml)}}})
			if e == nil {
				t.Fatal("invalid fresh registry accepted")
			}
			var out bytes.Buffer
			if e = p.WriteTo(&out); e != nil {
				t.Fatal(e)
			}
			if !bytes.Equal(out.Bytes(), source) {
				t.Fatal("refused registry dirtied package")
			}
		})
	}
}

func TestStandardDiagramMixedCaseMIMEIsPreserved(t *testing.T) {
	p := graphEditPackage(t)
	mime := "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
	plan, e := p.PlanGraphMutation(GraphMutation{Additions: []PartAddition{{Name: "word/diagram.xml", ContentType: mime, Data: []byte(`<data/>`)}}})
	if e != nil {
		t.Fatal(e)
	}
	if e = p.ApplyGraphPlan(plan); e != nil {
		t.Fatal(e)
	}
	graph, e := p.Graph()
	if e != nil {
		t.Fatal(e)
	}
	for _, part := range graph.Parts {
		if part.Name == "word/diagram.xml" {
			if part.ContentType != mime {
				t.Fatal("registered MIME spelling rewritten")
			}
			return
		}
	}
	t.Fatal("diagram part absent")
}
