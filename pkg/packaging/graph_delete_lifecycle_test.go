package packaging

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"testing"
)

func TestGraphDeletionLifecycleBatch(t *testing.T) {
	mustApply := func(p *Preserved, c GraphMutation) {
		t.Helper()
		plan, err := p.PlanGraphMutation(c)
		if err != nil {
			t.Fatal(err)
		}
		if err = p.ApplyGraphPlan(plan); err != nil {
			t.Fatal(err)
		}
	}
	t.Run("add replace delete cancels transient payload", func(t *testing.T) {
		p := graphEditPackage(t)
		source := serializeGraphEdit(t, p)
		mustApply(p, GraphMutation{Additions: []PartAddition{{"temp.bin", "application/octet-stream", []byte("first")}}})
		_, hash, _ := p.Part("temp.bin")
		mustApply(p, GraphMutation{Replacements: []Replacement{{"temp.bin", hash, []byte("second")}}})
		_, hash, _ = p.Part("temp.bin")
		mustApply(p, GraphMutation{Deletions: []PartDeletion{{"temp.bin", hash}}})
		if !bytes.Equal(source, serializeGraphEdit(t, p)) || len(p.Receipt().Changes) != 0 {
			t.Fatal("transient cancellation differs from source")
		}
		if _, _, err := p.Part("temp.bin"); err == nil {
			t.Fatal("transient payload still readable")
		}
	})
	t.Run("failed mixed plan leaves prior edits and held plan", func(t *testing.T) {
		p := graphEditPackage(t)
		_, hash, _ := p.Part("opaque.bin")
		if err := p.Replace([]Replacement{{"opaque.bin", hash, []byte("prior")}}); err != nil {
			t.Fatal(err)
		}
		_, hash, _ = p.Part("opaque.bin")
		good, err := p.PlanGraphMutation(GraphMutation{Deletions: []PartDeletion{{"opaque.bin", hash}}})
		if err != nil {
			t.Fatal(err)
		}
		before := serializeGraphEdit(t, p)
		bad, err := p.PlanGraphMutation(GraphMutation{Deletions: []PartDeletion{{"opaque.bin", hash}}, Replacements: []Replacement{{"word/document.xml", "stale", []byte(`<r>bad</r>`)}}})
		if err == nil || bad != nil {
			t.Fatal("failed private plan exposed")
		}
		if !bytes.Equal(before, serializeGraphEdit(t, p)) {
			t.Fatal("failed preview altered prior edits")
		}
		if err = p.ApplyGraphPlan(good); err != nil {
			t.Fatal("failed preview invalidated held plan", err)
		}
	})
	t.Run("deletion stales existing plans and reserves names", func(t *testing.T) {
		p := graphEditPackage(t)
		beforePlan, err := p.PlanGraphMutation(graphEditRequest())
		if err != nil {
			t.Fatal(err)
		}
		_, hash, _ := p.Part("opaque.bin")
		mustApply(p, GraphMutation{Deletions: []PartDeletion{{"opaque.bin", hash}}})
		err = p.ApplyGraphPlan(beforePlan)
		var refusal *Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "stale_target" {
			t.Fatal("plan not stale", err)
		}
		for _, name := range []string{"opaque.bin", "OPAQUE.bin"} {
			if _, err = p.PlanGraphMutation(GraphMutation{Additions: []PartAddition{{name, "application/octet-stream", nil}}}); err == nil {
				t.Fatal("deleted original identity reused")
			}
		}
		if err = p.Replace([]Replacement{{"opaque.bin", hash, nil}}); err == nil {
			t.Fatal("deleted payload replaced")
		}
	})
	t.Run("retarget removal and deletion compose without leaking data", func(t *testing.T) {
		p := graphEditPackage(t)
		_, hash, _ := p.Part("word/media/shared.bin")
		mustApply(p, GraphMutation{Additions: []PartAddition{{"word/media/new.bin", "application/octet-stream", []byte("private")}}, Retargets: []RelationshipRetarget{{"word/document.xml", "rId1", "word/media/new.bin"}}, Removals: []RelationshipRemoval{{"word/document.xml", "rId2"}}, Deletions: []PartDeletion{{"word/media/shared.bin", hash}}})
		g, err := p.Graph()
		if err != nil {
			t.Fatal(err)
		}
		for _, e := range g.Edges {
			if e.Source == "word/document.xml" && (e.ID != "rId1" || e.ResolvedPart != "word/media/new.bin") {
				t.Fatal("wrong surviving edge", e)
			}
		}
		dest := filepath.Join(t.TempDir(), "out.zip")
		receipt, err := p.SaveAs(dest)
		if err != nil {
			t.Fatal(err)
		}
		if receipt.Schema != 2 {
			t.Fatal(receipt)
		}
		out, err := os.ReadFile(dest)
		if err != nil {
			t.Fatal(err)
		}
		q, err := OpenPreserved(out, Limits{})
		if err != nil {
			t.Fatal(err)
		}
		if _, _, err = q.Part("word/media/shared.bin"); err == nil {
			t.Fatal("deleted member delivered")
		}
		if _, err = q.Graph(); err != nil {
			t.Fatal(err)
		}
	})
	t.Run("registry replacement and nonleaf deletion refuse", func(t *testing.T) {
		p := graphEditPackage(t)
		_, hash, _ := p.Part("word/document.xml")
		if _, err := p.PlanGraphMutation(GraphMutation{Deletions: []PartDeletion{{"word/document.xml", hash}}, Removals: []RelationshipRemoval{{"", "rId1"}, {"word/document.xml", "rId1"}, {"word/document.xml", "rId2"}}}); err == nil {
			t.Fatal("nonleaf deletion accepted")
		}
		reg := RelationshipsPathForPart("word/document.xml")
		b, h, _ := p.Part(reg)
		if _, err := p.PlanGraphMutation(GraphMutation{Retargets: []RelationshipRetarget{{"word/document.xml", "rId1", "opaque.bin"}}, Replacements: []Replacement{{reg, h, b}}}); err == nil {
			t.Fatal("registry patch overwritten by replacement")
		}
	})
}
