package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsSmartArtInspectionRecipes(t *testing.T) {
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
			Expected   []SmartArtInfo      `json:"expected"`
			ErrorCode  string              `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-smartart-inspection.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 12 {
		t.Fatal("SmartArt inspection inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			input := graphicsInput(t, base, c.Operations)
			s, e := OpenEditing(input, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsMembers(t, input)
			version := s.generation
			info, e := s.InspectSmartArt(r.SlidePart)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("SmartArt refusal %v want %s", e, c.ErrorCode)
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if !reflect.DeepEqual(info, c.Expected) {
					a, _ := json.MarshalIndent(info, "", " ")
					b, _ := json.MarshalIndent(c.Expected, "", " ")
					t.Fatalf("SmartArt records differ\ngot %s\nwant %s", a, b)
				}
				if len(info) > 0 {
					info[0].Name = "detached"
					if len(info[0].Roots) > 0 {
						info[0].Roots[0].Target = "detached"
					}
					if len(info[0].Parts) > 0 {
						info[0].Parts[0].PartName = "detached"
					}
					if len(info[0].Edges) > 0 {
						info[0].Edges[0].Target = "detached"
					}
				}
				next, er := s.InspectSmartArt(r.SlidePart)
				if er != nil || !reflect.DeepEqual(next, c.Expected) {
					t.Fatal("SmartArt detached readback", er)
				}
			}
			if s.generation != version {
				t.Fatal("inspection version")
			}
			path := filepath.Join(t.TempDir(), "smartart.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			if !reflect.DeepEqual(graphicsMembers(t, saved), before) {
				t.Fatal("SmartArt inspection custody")
			}
			if !bytes.Equal(input, graphicsInput(t, base, c.Operations)) {
				t.Fatal("source array modified")
			}
			if c.ErrorCode == "" {
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				next, er := reopened.InspectSmartArt(r.SlidePart)
				if er != nil || !reflect.DeepEqual(next, c.Expected) {
					t.Fatal("saved SmartArt readback", er)
				}
			}
		})
	}
}
func TestGraphicsSmartArtSealedSourceInspection(t *testing.T) {
	root, assets := graphicsRoot(t)
	var r struct {
		FixtureID string `json:"fixtureId"`
		ShapeID   uint32 `json:"shapeId"`
		Name      string `json:"name"`
		Expected  struct {
			PartNames []string `json:"partNames"`
		} `json:"expected"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-smartart-office-source.json"), &r); e != nil {
		t.Fatal(e)
	}
	input := graphicsReadAsset(t, root, assets, r.FixtureID)
	s, e := OpenEditing(input, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	info, e := s.InspectSmartArt("ppt/slides/slide1.xml")
	if e != nil {
		t.Fatal(e)
	}
	if len(info) != 1 || info[0].ShapeID != r.ShapeID || info[0].Name != r.Name || len(info[0].Roots) != 4 || len(info[0].Parts) != 5 {
		t.Fatalf("Office source graph %#v", info)
	}
	names := []string{}
	for _, p := range info[0].Parts {
		names = append(names, p.PartName)
	}
	if !reflect.DeepEqual(names, r.Expected.PartNames) {
		t.Fatal("source closure parts")
	}
	if len(info[0].Edges) != 1 || info[0].Edges[0].Owner != "ppt/slides/slide1.xml" || info[0].Edges[0].RelationshipID != "rId6" {
		t.Fatal("slide-owned drawing metadata edge")
	}
}
