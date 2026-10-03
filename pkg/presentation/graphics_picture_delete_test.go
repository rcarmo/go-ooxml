package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsPictureDeletionRecipes(t *testing.T) {
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
				Options json.RawMessage `json:"options"`
			} `json:"request"`
			Expected  *PictureDeleteReceipt `json:"expected"`
			ErrorCode string                `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-delete.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 17 {
		t.Fatal("delete inventory")
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
			var options PictureDeleteOptions
			d := json.NewDecoder(bytes.NewReader(c.Request.Options))
			d.DisallowUnknownFields()
			e = d.Decode(&options)
			var receipt PictureDeleteReceipt
			if e != nil {
				e = graphicsUnsupported("native deletion options decode")
			} else {
				receipt, e = s.DeletePicture(r.SlidePart, c.Request.ShapeID, options)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("delete refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("refusal generation")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if !reflect.DeepEqual(receipt, *c.Expected) || s.generation != version+1 {
					t.Fatalf("delete receipt %#v want %#v", receipt, c.Expected)
				}
			}
			path := filepath.Join(t.TempDir(), "deleted.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			after := graphicsMembers(t, saved)
			removed := map[string]bool{}
			for _, n := range receipt.RemovedMedia {
				removed[n] = true
			}
			if len(after) != len(before)-len(removed) {
				t.Fatal("delete membership")
			}
			for n, b := range before {
				if removed[n] {
					if _, exists := after[n]; exists {
						t.Fatal("removed media remains")
					}
					continue
				}
				allowed := c.ErrorCode == "" && (n == r.SlidePart || n == packaging.RelationshipsPathForPart(r.SlidePart) || n == packaging.ContentTypesPath)
				if !allowed && !bytes.Equal(b, after[n]) {
					t.Fatal("deletion custody", n)
				}
			}
			if c.ErrorCode == "" {
				doc, er := losslessxml.Parse(before[r.SlidePart])
				if er != nil {
					t.Fatal(er)
				}
				var selected losslessxml.Element
				for _, n := range doc.Elements() {
					if n.Name() != name(packaging.NSPresentationML, "pic") {
						continue
					}
					id, _, _, er := graphicsIdentity(doc, n, "nvPicPr")
					if er != nil {
						t.Fatal(er)
					}
					if id == c.Request.ShapeID {
						selected = n
					}
				}
				a, z := selected.SourceRange()
				expected := append(append([]byte{}, before[r.SlidePart][:a]...), before[r.SlidePart][z:]...)
				if !bytes.Equal(expected, after[r.SlidePart]) {
					t.Fatal("deleted slide not exact span removal")
				}
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				old, er := s.InspectPictures(r.SlidePart)
				if er != nil {
					t.Fatal(er)
				}
				next, er := reopened.InspectPictures(r.SlidePart)
				if er != nil || !reflect.DeepEqual(old, next) {
					t.Fatal("deletion saved records", er)
				}
				if _, er = s.pkg.Graph(); er != nil {
					t.Fatal("deleted graph invalid", er)
				}
				graphicsWriteOutput(t, root, "picture-delete-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
