package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"reflect"
	"testing"
)

func TestGraphicsOutlineRecipes(t *testing.T) {
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
				ShapeID uint32          `json:"shapeId"`
				Patch   json.RawMessage `json:"patch"`
			} `json:"request"`
			Expected struct {
				Before  OutlineStyle `json:"before"`
				Outline OutlineStyle `json:"outline"`
				Changed int          `json:"changed"`
			} `json:"expected"`
			ErrorCode string `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-outlines.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 21 {
		t.Fatal("outline inventory")
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
				got, er := s.GetOutlineStyle(r.SlidePart, c.Request.ShapeID)
				if er != nil || !reflect.DeepEqual(got, c.Expected.Before) {
					t.Fatal("outline before", got, er)
				}
			}
			var patch OutlinePatch
			d := json.NewDecoder(bytes.NewReader(c.Request.Patch))
			d.DisallowUnknownFields()
			e = d.Decode(&patch)
			changed := 0
			if e != nil {
				e = editRefusal("PPTX_OUTLINE_UNSUPPORTED", "native outline decoding")
			} else {
				changed, e = s.PatchOutlineStyle(r.SlidePart, c.Request.ShapeID, patch)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("outline refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, s)) {
					t.Fatal("outline refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if changed != c.Expected.Changed || s.generation != version+uint64(changed) {
				t.Fatal("outline changed/version")
			}
			got, e := s.GetOutlineStyle(r.SlidePart, c.Request.ShapeID)
			if e != nil || !reflect.DeepEqual(got, c.Expected.Outline) {
				t.Fatal("outline readback", got, e)
			}
			if got.HeadEnd != nil {
				got.HeadEnd.Type = "none"
			}
			read, e := s.GetOutlineStyle(r.SlidePart, c.Request.ShapeID)
			if e != nil || !reflect.DeepEqual(read, c.Expected.Outline) {
				t.Fatal("outline aliasing")
			}
			after := graphicsSessionMembers(t, s)
			graphicsAssertStyleCustody(t, before[r.SlidePart], after[r.SlidePart], c.Request.ShapeID, "outline")
			if len(after) != len(before) {
				t.Fatal("outline membership")
			}
			for n, b := range before {
				if n != r.SlidePart || changed == 0 {
					if !bytes.Equal(b, after[n]) {
						t.Fatal("outline custody", n)
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
			value, e := reopened.GetOutlineStyle(r.SlidePart, c.Request.ShapeID)
			if e != nil || !reflect.DeepEqual(value, c.Expected.Outline) {
				t.Fatal("outline reopen", e)
			}
			graphicsWriteOutput(t, root, "outlines-"+c.CaseID+".pptx", out.Bytes())
		})
	}
}
