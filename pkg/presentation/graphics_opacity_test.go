package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"reflect"
	"testing"
)

func TestGraphicsOpacityRecipes(t *testing.T) {
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
				Profile string          `json:"profile"`
				ShapeID uint32          `json:"shapeId"`
				Value   json.RawMessage `json:"value"`
			} `json:"request"`
			Expected struct {
				Before  int64 `json:"before"`
				Value   int64 `json:"value"`
				Changed int   `json:"changed"`
			} `json:"expected"`
			ErrorCode string `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-opacity.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 24 {
		t.Fatal("alpha inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			s, e := OpenEditing(graphicsInput(t, base, c.Operations), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsSessionMembers(t, s)
			version := s.generation
			get := s.GetShapeOpacity
			set := s.SetShapeOpacity
			if c.Request.Profile == "picture" {
				get = s.GetPictureTransparency
				set = s.SetPictureTransparency
			}
			if c.ErrorCode == "" {
				v, er := get(r.SlidePart, c.Request.ShapeID)
				if er != nil || v != c.Expected.Before {
					t.Fatal("alpha before", v, er)
				}
			}
			var value int64
			e = json.Unmarshal(c.Request.Value, &value)
			changed := 0
			if e != nil {
				e = editRefusal("PPTX_OPACITY_UNSUPPORTED", "native alpha decoding")
			} else {
				changed, e = set(r.SlidePart, c.Request.ShapeID, value)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("alpha refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, s)) {
					t.Fatal("alpha refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if changed != c.Expected.Changed || s.generation != version+uint64(changed) {
				t.Fatal("alpha changed/version")
			}
			got, e := get(r.SlidePart, c.Request.ShapeID)
			if e != nil || got != c.Expected.Value {
				t.Fatal("alpha readback", got, e)
			}
			after := graphicsSessionMembers(t, s)
			profile := "shape-alpha"
			if c.Request.Profile == "picture" {
				profile = "picture-alpha"
			}
			graphicsAssertStyleCustody(t, before[r.SlidePart], after[r.SlidePart], c.Request.ShapeID, profile)
			if len(after) != len(before) {
				t.Fatal("alpha membership")
			}
			for n, b := range before {
				if n != r.SlidePart || changed == 0 {
					if !bytes.Equal(b, after[n]) {
						t.Fatal("alpha custody", n)
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
			get = reopened.GetShapeOpacity
			if c.Request.Profile == "picture" {
				get = reopened.GetPictureTransparency
			}
			v, e := get(r.SlidePart, c.Request.ShapeID)
			if e != nil || v != c.Expected.Value {
				t.Fatal("alpha reopen", e)
			}
			graphicsWriteOutput(t, root, "opacity-"+c.CaseID+".pptx", out.Bytes())
		})
	}
}
