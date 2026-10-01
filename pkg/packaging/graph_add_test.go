package packaging

import (
	"bytes"
	"errors"
	"testing"
)

func TestGraphAdditionAndEffectiveTypeFamily(t *testing.T) {
	source := graphEditPackage(t)
	before := serializeGraphEdit(t, source)
	change := GraphMutation{Additions: []PartAddition{{Name: "custom/data.bin", ContentType: "application/octet-stream", Data: []byte{7, 8, 9}}}, Relationships: []RelationshipAddition{{Source: "", ID: "rIdData", Type: "urn:test/data", TargetPart: "custom/data.bin"}}}
	plan, err := source.PlanGraphMutation(change)
	if err != nil {
		t.Fatal(err)
	}
	change.Additions[0].Data[0] = 255
	if err := source.ApplyGraphPlan(plan); err != nil {
		t.Fatal(err)
	}
	graph, err := source.Graph()
	if err != nil {
		t.Fatal(err)
	}
	foundPart, foundEdge := false, false
	for _, part := range graph.Parts {
		if part.Name == "custom/data.bin" && part.ContentType == "application/octet-stream" {
			foundPart = true
		}
	}
	for _, edge := range graph.Edges {
		if edge.Source == "" && edge.ID == "rIdData" && edge.Type == "urn:test/data" && edge.ResolvedPart == "custom/data.bin" {
			foundEdge = true
		}
	}
	if !foundPart || !foundEdge {
		t.Fatalf("graph add part=%v edge=%v", foundPart, foundEdge)
	}
	payload, hash, err := source.Part("custom/data.bin")
	if err != nil || !bytes.Equal(payload, []byte{7, 8, 9}) {
		t.Fatalf("payload: %x %v", payload, err)
	}
	baseline := serializeGraphEdit(t, source)
	if rejected, err := source.PlanGraphMutation(GraphMutation{Deletions: []PartDeletion{{Name: "custom/data.bin", ExpectedSHA256: hash}}}); rejected != nil || err == nil {
		t.Fatal("referenced part deleted")
	} else {
		var refused *Refusal
		if !errors.As(err, &refused) || refused.Kind != "opc-part-referenced" {
			t.Fatalf("referenced deletion reason: %v", err)
		}
	}
	if !bytes.Equal(baseline, serializeGraphEdit(t, source)) {
		t.Fatal("refusal changed archive")
	}
	// A failed graph operation must not consume the session; an existing part
	// handle and unrelated parts remain live and immutable.
	retained, latest, err := source.Part("custom/data.bin")
	if err != nil || latest != hash || !bytes.Equal(retained, []byte{7, 8, 9}) {
		t.Fatalf("refusal changed held part: %v", err)
	}
	for _, kind := range []GraphMutation{
		{Relationships: []RelationshipAddition{{ID: "rIdData", Type: "urn:test/duplicate", TargetPart: "custom/data.bin"}}},
		{Relationships: []RelationshipAddition{{ID: "rIdMissing", Type: "urn:test/data", TargetPart: "missing.bin"}}},
		{ContentTypes: []ContentTypeChange{{Part: "missing.bin", ContentType: "application/octet-stream"}}},
	} {
		if plan, err := source.PlanGraphMutation(kind); plan != nil || err == nil {
			t.Fatalf("graph refusal delivered plan: %v", err)
		}
		if !bytes.Equal(baseline, serializeGraphEdit(t, source)) {
			t.Fatal("graph refusal changed current archive")
		}
	}
	detach, err := source.PlanGraphMutation(GraphMutation{Removals: []RelationshipRemoval{{Source: "", ID: "rIdData"}}, Deletions: []PartDeletion{{Name: "custom/data.bin", ExpectedSHA256: hash}}})
	if err != nil {
		t.Fatal(err)
	}
	if err := source.ApplyGraphPlan(detach); err != nil {
		t.Fatal(err)
	}
	if _, _, err := source.Part("custom/data.bin"); err == nil {
		t.Fatal("deleted part retained")
	}
	if _, err := source.Graph(); err != nil {
		t.Fatal(err)
	}
	if len(before) == 0 {
		t.Fatal("empty original archive")
	}
}

func TestGraphTypeOnlyChangeAndRefusals(t *testing.T) {
	first := graphEditPackage(t)
	second := graphEditPackage(t)
	original, _, err := first.Part("word/document.xml")
	if err != nil {
		t.Fatal(err)
	}
	plan, err := second.PlanGraphMutation(GraphMutation{ContentTypes: []ContentTypeChange{{Part: "word/document.xml", ContentType: "application/vnd.test.document+xml"}}})
	if err != nil {
		t.Fatal(err)
	}
	if err = second.ApplyGraphPlan(plan); err != nil {
		t.Fatal(err)
	}
	edited, _, err := second.Part("word/document.xml")
	if err != nil || !bytes.Equal(edited, original) {
		t.Fatal("type-only edit changed payload")
	}
	graph, err := second.Graph()
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, part := range graph.Parts {
		if part.Name == "word/document.xml" && part.ContentType == "application/vnd.test.document+xml" {
			found = true
		}
	}
	if !found {
		t.Fatal("type not effective")
	}
	if p, err := first.PlanGraphMutation(GraphMutation{Relationships: []RelationshipAddition{{ID: "rId1", Type: "urn:test/data", TargetPart: "word/document.xml"}}}); p != nil || err == nil {
		t.Fatal("duplicate edge accepted")
	}
	if p, err := first.PlanGraphMutation(GraphMutation{Relationships: []RelationshipAddition{{ID: "new", Type: "urn:test/data", TargetPart: "missing.bin"}}}); p != nil || err == nil {
		t.Fatal("missing edge target accepted")
	}
	var refused *Refusal
	if p, err := first.PlanGraphMutation(GraphMutation{ContentTypes: []ContentTypeChange{{Part: "missing.bin", ContentType: "application/octet-stream"}}}); p != nil || !errors.As(err, &refused) {
		t.Fatal("missing MIME target accepted")
	}
}
