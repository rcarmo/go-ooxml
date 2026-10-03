package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"math"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsGroupTransformRecipes(t *testing.T) {
	root, assets := graphicsRoot(t)
	var r struct {
		Feature       string `json:"feature"`
		Contract      string `json:"contract"`
		BaseFixtureID string `json:"baseFixtureId"`
		SlidePart     string `json:"slidePart"`
		Tolerance     struct {
			Absolute float64 `json:"absoluteEMUs"`
			Relative float64 `json:"relative"`
		} `json:"tolerance"`
		Cases []struct {
			ScenarioID string              `json:"scenarioId"`
			CaseID     string              `json:"caseId"`
			Operations []graphicsOperation `json:"operations"`
			Request    struct {
				ShapeID uint32          `json:"shapeId"`
				Patch   json.RawMessage `json:"patch"`
			} `json:"request"`
			Expected struct {
				Before    GroupTransform `json:"before"`
				Transform GroupTransform `json:"transform"`
				Changed   int            `json:"changed"`
			} `json:"expected"`
			Mapping *struct {
				Chain  []GroupTransform `json:"transforms"`
				Input  GroupPoint       `json:"point"`
				Output GroupPoint       `json:"expected"`
			} `json:"mapping"`
			ErrorCode string `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-group-transform.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 21 {
		t.Fatal("group frame inventory")
	}
	closePoint := func(t *testing.T, got, want GroupPoint) {
		t.Helper()
		for _, v := range []struct{ a, b float64 }{{got.X, want.X}, {got.Y, want.Y}} {
			if math.Abs(v.a-v.b) > r.Tolerance.Absolute+r.Tolerance.Relative*math.Abs(v.b) {
				t.Fatalf("mapping %#v want %#v", got, want)
			}
		}
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			if c.Mapping != nil {
				got, e := MapGroupPointChain(c.Mapping.Chain, c.Mapping.Input)
				if e != nil {
					t.Fatal(e)
				}
				closePoint(t, got, c.Mapping.Output)
				inverse, e := UnmapGroupPointChain(c.Mapping.Chain, got)
				if e != nil {
					t.Fatal(e)
				}
				closePoint(t, inverse, c.Mapping.Input)
				return
			}
			input := graphicsInput(t, base, c.Operations)
			s, e := OpenEditing(input, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsMembers(t, input)
			version := s.generation
			var patch GroupTransformPatch
			d := json.NewDecoder(bytes.NewReader(c.Request.Patch))
			d.DisallowUnknownFields()
			e = d.Decode(&patch)
			changed := 0
			if e != nil {
				e = graphicsGroupUnsupported("native group patch decode")
			} else {
				if c.ErrorCode == "" {
					got, er := s.GetGroupTransform(r.SlidePart, c.Request.ShapeID)
					if er != nil || got != c.Expected.Before {
						t.Fatal("group frame before", got, er)
					}
				}
				changed, e = s.PatchGroupTransform(r.SlidePart, c.Request.ShapeID, patch)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("frame refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("refusal generation")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if changed != c.Expected.Changed || s.generation != version+uint64(changed) {
					t.Fatal("group changed/version")
				}
				got, er := s.GetGroupTransform(r.SlidePart, c.Request.ShapeID)
				if er != nil || got != c.Expected.Transform {
					t.Fatal("group frame literal readback", got, er)
				}
			}
			path := filepath.Join(t.TempDir(), "frame.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			after := graphicsMembers(t, saved)
			if len(after) != len(before) {
				t.Fatal("frame membership")
			}
			for n, b := range before {
				if c.ErrorCode != "" || n != r.SlidePart || changed == 0 {
					if !bytes.Equal(after[n], b) {
						t.Fatal("frame unrelated bytes", n)
					}
				}
			}
			if c.ErrorCode == "" {
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				got, er := reopened.GetGroupTransform(r.SlidePart, c.Request.ShapeID)
				if er != nil || got != c.Expected.Transform {
					t.Fatal("saved frame readback", er)
				}
				a, er := s.InspectPictures(r.SlidePart)
				if er != nil {
					t.Fatal(er)
				}
				b, er := reopened.InspectPictures(r.SlidePart)
				if er != nil || !reflect.DeepEqual(a, b) {
					t.Fatal("saved children metadata", er)
				}
				graphicsWriteOutput(t, root, "group-transform-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
