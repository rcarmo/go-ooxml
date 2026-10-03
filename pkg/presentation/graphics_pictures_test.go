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

func TestGraphicsPictureInspectionRecipes(t *testing.T) {
	root, assets := graphicsRoot(t)
	const ledger = "ledgers/pptx-picture-inspection.json"
	var recipe graphicsPictureRecipe
	if err := json.Unmarshal(graphicsReadAsset(t, root, assets, ledger), &recipe); err != nil {
		t.Fatal(err)
	}
	graphicsReadAsset(t, root, assets, recipe.Feature)
	graphicsReadAsset(t, root, assets, recipe.Contract)
	base := graphicsReadAsset(t, root, assets, recipe.BaseFixtureID)
	if len(recipe.Cases) != 12 {
		t.Fatal("picture recipe case inventory")
	}
	for _, c := range recipe.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			input := graphicsInput(t, base, c.Operations)
			original := bytes.Clone(input)
			before := graphicsMembers(t, input)
			session, err := OpenEditing(input, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			version := session.generation
			got, err := session.InspectPictures(recipe.SlidePart)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(err, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("refusal %v; want %s", err, c.ErrorCode)
				}
			} else {
				if err != nil {
					t.Fatal(err)
				}
				if !reflect.DeepEqual(got, c.Expected) {
					actual, _ := json.MarshalIndent(got, "", "  ")
					expected, _ := json.MarshalIndent(c.Expected, "", "  ")
					t.Fatalf("picture records differ\ngot %s\nwant %s", actual, expected)
				}
				if len(got) > 0 {
					got[0].Name = "detached"
					if got[0].Transform != nil {
						got[0].Transform.X = -99
					}
					if got[0].Embedded != nil {
						got[0].Embedded.Target = "detached"
					}
					if len(got[0].Groups) > 0 {
						got[0].Groups[0].Name = "detached"
						if got[0].Groups[0].Transform != nil {
							got[0].Groups[0].Transform.ChildX = -99
						}
					}
				}
				next, err := session.InspectPictures(recipe.SlidePart)
				if err != nil || !reflect.DeepEqual(next, c.Expected) {
					t.Fatal("inspection results aliased package state", err)
				}
			}
			if session.generation != version || !bytes.Equal(input, original) {
				t.Fatal("inspection changed version or input")
			}
			path := filepath.Join(t.TempDir(), "inspection.pptx")
			if _, err = session.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(graphicsMembers(t, saved), before) {
				t.Fatal("inspection changed member payloads or membership")
			}
			reopened, err := OpenEditing(saved, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			next, err := reopened.InspectPictures(recipe.SlidePart)
			if c.ErrorCode == "" && (err != nil || !reflect.DeepEqual(next, c.Expected)) {
				t.Fatal("saved readback differs", err)
			}
		})
	}
}
