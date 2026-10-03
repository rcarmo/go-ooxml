package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"reflect"
	"testing"
)

func TestGraphicsZOrderRecipes(t *testing.T) {
	root, assets := graphicsRoot(t)
	var r struct {
		Feature       string `json:"feature"`
		Contract      string `json:"contract"`
		BaseFixtureID string `json:"baseFixtureId"`
		SlidePart     string `json:"slidePart"`
		Cases         []struct {
			ScenarioID string              `json:"scenarioId"`
			CaseID     string              `json:"caseId"`
			Operations []graphicsOperation `json:"operations"`
			Request    struct {
				Order   []uint32 `json:"order"`
				GroupID *uint32  `json:"groupId"`
			} `json:"request"`
			Expected struct {
				Before  []uint32 `json:"before"`
				Order   []uint32 `json:"order"`
				Changed int      `json:"changed"`
			} `json:"expected"`
			ErrorCode string `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-z-order.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 15 {
		t.Fatal("order inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			s, e := OpenEditing(graphicsInput(t, base, c.Operations), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsSessionMembers(t, s)
			version := s.generation
			if c.ErrorCode == "" {
				order, er := s.GetShapeOrder(r.SlidePart, c.Request.GroupID)
				if er != nil || !reflect.DeepEqual(order, c.Expected.Before) {
					t.Fatal("order before", order, er)
				}
			}
			changed, e := s.ReorderShapes(r.SlidePart, c.Request.Order, c.Request.GroupID)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("order refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, s)) {
					t.Fatal("order refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if changed != c.Expected.Changed || s.generation != version+uint64(changed) {
				t.Fatal("order changed/version")
			}
			order, e := s.GetShapeOrder(r.SlidePart, c.Request.GroupID)
			if e != nil || !reflect.DeepEqual(order, c.Expected.Order) {
				t.Fatal("order literal readback", order, e)
			}
			after := graphicsSessionMembers(t, s)
			if len(after) != len(before) {
				t.Fatal("order membership")
			}
			for n, b := range before {
				if n != r.SlidePart || changed == 0 {
					if !bytes.Equal(b, after[n]) {
						t.Fatal("order payload custody", n)
					}
				}
			}
			if _, e = s.ReorderShapes(r.SlidePart, c.Expected.Before, c.Request.GroupID); e != nil {
				t.Fatal(e)
			}
			restored := graphicsSessionMembers(t, s)
			if !reflect.DeepEqual(restored, before) {
				t.Fatal("inverse permutation did not recover original spans")
			}
			var out bytes.Buffer
			// Reapply requested order for persistence/application checks.
			if _, e = s.ReorderShapes(r.SlidePart, c.Request.Order, c.Request.GroupID); e != nil {
				t.Fatal(e)
			}
			if e = s.pkg.WriteTo(&out); e != nil {
				t.Fatal(e)
			}
			reopened, e := OpenEditing(out.Bytes(), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			read, e := reopened.GetShapeOrder(r.SlidePart, c.Request.GroupID)
			if e != nil || !reflect.DeepEqual(read, c.Expected.Order) {
				t.Fatal("saved order", e)
			}
			graphicsWriteOutput(t, root, "z-order-"+c.CaseID+".pptx", out.Bytes())
		})
	}
}
