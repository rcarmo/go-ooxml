package packaging

import (
	"bytes"
	"testing"
)

func TestRetargetURIAndChainedPlans(t *testing.T) {
	for _, source := range []string{"", "a.xml", "a/b/c.xml", "a/media/source.xml"} {
		for _, target := range []string{"one.bin", "a/new.bin", "a/media/new & 雪.bin", "other/deep/a'b.bin"} {
			for _, original := range []string{"old.bin", "../old.bin", "/old.bin"} {
				uri := retargetURI(source, target, original)
				got, err := resolveGraphTarget(source, uri)
				if err != nil || got != target {
					t.Fatalf("%q %q => %q: %q %v", source, target, uri, got, err)
				}
			}
		}
	}
	t.Run("multiple additions and patches accumulate", func(t *testing.T) {
		p := graphEditPackage(t)
		for _, name := range []string{"word/media/a & 雪.bin", "word/media/b.bin"} {
			plan, err := p.PlanGraphMutation(GraphMutation{Additions: []PartAddition{{Name: name, ContentType: "application/octet-stream", Data: []byte(name)}}, Retargets: []RelationshipRetarget{{Source: "word/document.xml", ID: "rId1", TargetPart: name}}})
			if err != nil {
				t.Fatal(err)
			}
			if err = p.ApplyGraphPlan(plan); err != nil {
				t.Fatal(err)
			}
		}
		out := serializeGraphEdit(t, p)
		q, err := OpenPreserved(out, Limits{})
		if err != nil {
			t.Fatal(err)
		}
		if _, err = q.Graph(); err != nil {
			t.Fatal(err)
		}
		for _, name := range []string{"word/media/a & 雪.bin", "word/media/b.bin", "word/media/shared.bin"} {
			if _, _, err = q.Part(name); err != nil {
				t.Fatal(err)
			}
		}
		if len(p.Receipt().Changes) != 4 {
			t.Fatal("additions not accumulated")
		}
	})
	t.Run("existing target retarget-only and exact restore", func(t *testing.T) {
		p := graphEditPackage(t)
		original := serializeGraphEdit(t, p)
		plan, err := p.PlanGraphMutation(GraphMutation{Retargets: []RelationshipRetarget{{"word/document.xml", "rId1", "opaque.bin"}}})
		if err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
		if p.Receipt().Schema != 1 || len(p.Receipt().Changes) != 1 {
			t.Fatal("retarget-only receipt")
		}
		plan, err = p.PlanGraphMutation(GraphMutation{Retargets: []RelationshipRetarget{{"word/document.xml", "rId1", "word/media/shared.bin"}}})
		if err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(original, serializeGraphEdit(t, p)) {
			t.Fatal("restore did not return exact original")
		}
	})
}
