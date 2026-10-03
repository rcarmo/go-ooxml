package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"reflect"
	"testing"
)

func TestGraphicsGradientRecipes(t *testing.T) {
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
				ShapeID  uint32          `json:"shapeId"`
				Gradient json.RawMessage `json:"gradient"`
			} `json:"request"`
			Expected struct {
				Before   *LinearGradient `json:"before"`
				Gradient LinearGradient  `json:"gradient"`
				Changed  int             `json:"changed"`
			} `json:"expected"`
			ErrorCode string `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-gradients.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 18 {
		t.Fatal("gradient inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			s, e := OpenEditing(graphicsInput(t, base, c.Operations), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsSessionMembers(t, s)
			version := s.generation
			var g LinearGradient
			d := json.NewDecoder(bytes.NewReader(c.Request.Gradient))
			d.DisallowUnknownFields()
			e = d.Decode(&g)
			changed := 0
			if e != nil {
				e = editRefusal("PPTX_GRADIENT_UNSUPPORTED", "native gradient decoding")
			} else {
				if c.ErrorCode == "" {
					got, er := s.GetLinearGradient(r.SlidePart, c.Request.ShapeID)
					if er != nil || !reflect.DeepEqual(got, c.Expected.Before) {
						t.Fatal("gradient before", got, er)
					}
				}
				changed, e = s.SetLinearGradient(r.SlidePart, c.Request.ShapeID, g)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("gradient refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, s)) {
					t.Fatal("gradient refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if changed != c.Expected.Changed || s.generation != version+uint64(changed) {
				t.Fatal("gradient changed/version")
			}
			got, e := s.GetLinearGradient(r.SlidePart, c.Request.ShapeID)
			if e != nil || got == nil || !reflect.DeepEqual(*got, c.Expected.Gradient) {
				t.Fatal("gradient literal readback", got, e)
			}
			after := graphicsSessionMembers(t, s)
			for n, b := range before {
				if n != r.SlidePart || changed == 0 {
					if !bytes.Equal(b, after[n]) {
						t.Fatal("gradient custody", n)
					}
				}
			}
			var out bytes.Buffer
			if e = s.pkg.WriteTo(&out); e != nil {
				t.Fatal(e)
			}
			reopened, e := OpenEditing(out.Bytes(), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			read, e := reopened.GetLinearGradient(r.SlidePart, c.Request.ShapeID)
			if e != nil || !reflect.DeepEqual(read, got) {
				t.Fatal("gradient reopen", e)
			}
			graphicsWriteOutput(t, root, "gradients-"+c.CaseID+".pptx", out.Bytes())
		})
	}
}
