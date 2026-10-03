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

func TestGraphicsPicturePlacementRecipes(t *testing.T) {
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
				Payload struct {
					FixtureID string `json:"fixtureId"`
					Part      string `json:"part"`
				} `json:"payload"`
				Box     json.RawMessage `json:"box"`
				Options json.RawMessage `json:"options"`
			} `json:"request"`
			RequestPatch struct {
				Box     map[string]json.RawMessage `json:"box"`
				Options map[string]json.RawMessage `json:"options"`
			} `json:"requestPatch"`
			Expected  *FittedPictureReceipt `json:"expected"`
			ErrorCode string                `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-placement.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 16 {
		t.Fatal("placement inventory")
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
			payload := graphicsMembers(t, graphicsReadAsset(t, root, assets, c.Request.Payload.FixtureID))[c.Request.Payload.Part]
			var rawBox, rawOptions map[string]json.RawMessage
			if e = json.Unmarshal(c.Request.Box, &rawBox); e != nil {
				t.Fatal(e)
			}
			if e = json.Unmarshal(c.Request.Options, &rawOptions); e != nil {
				t.Fatal(e)
			}
			for k, v := range c.RequestPatch.Box {
				rawBox[k] = v
			}
			for k, v := range c.RequestPatch.Options {
				rawOptions[k] = v
			}
			c.Request.Box, _ = json.Marshal(rawBox)
			c.Request.Options, _ = json.Marshal(rawOptions)
			var box PictureGeometry
			var options FittedPictureOptions
			d := json.NewDecoder(bytes.NewReader(c.Request.Box))
			d.DisallowUnknownFields()
			e = d.Decode(&box)
			if e == nil {
				d = json.NewDecoder(bytes.NewReader(c.Request.Options))
				d.DisallowUnknownFields()
				e = d.Decode(&options)
			}
			var receipt FittedPictureReceipt
			if e != nil {
				e = graphicsUnsupported("native fitted request decoding")
			} else {
				receipt, e = s.AddFittedPicture(r.SlidePart, payload, box, options)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("refusal version")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				c.Expected.PartName = r.SlidePart
				if !reflect.DeepEqual(receipt, *c.Expected) {
					t.Fatalf("placement %#v want %#v", receipt, c.Expected)
				}
				pure, er := CalculatePicturePlacement(box, options.Intrinsic, options.Fit)
				if er != nil || pure.Geometry != c.Expected.Geometry || pure.Crop != c.Expected.Crop {
					t.Fatal("pure placement differs", pure, er)
				}
			}
			path := filepath.Join(t.TempDir(), "fitted.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			after := graphicsMembers(t, saved)
			extra := 0
			if c.ErrorCode == "" {
				extra = 1
			}
			if len(after) != len(before)+extra {
				t.Fatal("placement membership")
			}
			for n, b := range before {
				allowed := c.ErrorCode == "" && (n == r.SlidePart || n == packaging.RelationshipsPathForPart(r.SlidePart) || n == packaging.ContentTypesPath)
				if !allowed && !bytes.Equal(b, after[n]) {
					t.Fatal("placement custody", n)
				}
			}
			if c.ErrorCode == "" {
				if !bytes.Equal(after[receipt.MediaPart], payload) {
					t.Fatal("fit media bytes")
				}
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				pics, er := reopened.InspectPictures(r.SlidePart)
				if er != nil {
					t.Fatal(er)
				}
				found := false
				for _, p := range pics {
					if p.ShapeID == receipt.ShapeID {
						found = true
						if p.Transform == nil || p.Transform.X != receipt.Geometry.X || p.Transform.Y != receipt.Geometry.Y || p.Transform.Width != receipt.Geometry.Width || p.Transform.Height != receipt.Geometry.Height || p.Crop != receipt.Crop {
							t.Fatal("saved fit values", p)
						}
					}
				}
				if !found {
					t.Fatal("fitted picture absent")
				}
				graphicsWriteOutput(t, root, "picture-placement-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
