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

func TestGraphicsPictureMetadataRecipes(t *testing.T) {
	root, assets := graphicsRoot(t)
	for _, family := range []string{"crop", "transform"} {
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
					Crop    json.RawMessage `json:"crop"`
					Patch   json.RawMessage `json:"patch"`
				} `json:"request"`
				Expected struct {
					Before    PictureCrop      `json:"beforeCrop"`
					Crop      PictureCrop      `json:"crop"`
					Transform PictureTransform `json:"transform"`
					Changed   int              `json:"changed"`
				} `json:"expected"`
				ErrorCode string `json:"errorCode"`
			} `json:"cases"`
		}
		if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-"+family+".json"), &r); e != nil {
			t.Fatal(e)
		}
		graphicsReadAsset(t, root, assets, r.Feature)
		graphicsReadAsset(t, root, assets, r.Contract)
		base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
		wantCount := 17
		if family == "transform" {
			wantCount = 18
		}
		if len(r.Cases) != wantCount {
			t.Fatal("metadata case inventory")
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
				changed := 0
				if family == "crop" {
					var crop PictureCrop
					decoder := json.NewDecoder(bytes.NewReader(c.Request.Crop))
					decoder.DisallowUnknownFields()
					e = decoder.Decode(&crop)
					var fields map[string]json.RawMessage
					if er := json.Unmarshal(c.Request.Crop, &fields); er == nil && len(fields) != 4 {
						e = errors.New("crop requires four sides")
					}
					if e != nil {
						e = graphicsUnsupported("native crop decoding")
					} else {
						if c.ErrorCode == "" {
							got, er := s.GetPictureCrop(r.SlidePart, c.Request.ShapeID)
							if er != nil || got != c.Expected.Before {
								t.Fatal("crop before", got, er)
							}
						}
						changed, e = s.SetPictureCrop(r.SlidePart, c.Request.ShapeID, crop)
					}
				} else {
					var patch PictureTransformPatch
					decoder := json.NewDecoder(bytes.NewReader(c.Request.Patch))
					decoder.DisallowUnknownFields()
					e = decoder.Decode(&patch)
					if e != nil {
						e = graphicsUnsupported("native orientation decoding")
					} else {
						changed, e = s.PatchPictureTransform(r.SlidePart, c.Request.ShapeID, patch)
					}
				}
				if c.ErrorCode != "" {
					var refusal *packaging.Refusal
					if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
						t.Fatalf("refusal %v want %s", e, c.ErrorCode)
					}
					if s.generation != version {
						t.Fatal("refusal changed generation")
					}
				} else {
					if e != nil {
						t.Fatal(e)
					}
					if changed != c.Expected.Changed || s.generation != version+uint64(changed) {
						t.Fatal("changed/version")
					}
					if family == "crop" {
						got, er := s.GetPictureCrop(r.SlidePart, c.Request.ShapeID)
						if er != nil || got != c.Expected.Crop {
							t.Fatal("crop readback", got, er)
						}
					} else {
						pics, er := s.InspectPictures(r.SlidePart)
						if er != nil {
							t.Fatal(er)
						}
						found := false
						for _, p := range pics {
							if p.ShapeID == c.Request.ShapeID {
								found = true
								if p.Transform == nil || *p.Transform != c.Expected.Transform {
									t.Fatal("transform literal readback", p.Transform)
								}
							}
						}
						if !found {
							t.Fatal("picture absent")
						}
					}
				}
				path := filepath.Join(t.TempDir(), "metadata.pptx")
				if _, e = s.SaveAs(path); e != nil {
					t.Fatal(e)
				}
				saved, e := os.ReadFile(path)
				if e != nil {
					t.Fatal(e)
				}
				after := graphicsMembers(t, saved)
				if len(after) != len(before) {
					t.Fatal("metadata membership")
				}
				for n, b := range before {
					if c.ErrorCode != "" || n != r.SlidePart || changed == 0 {
						if !bytes.Equal(after[n], b) {
							t.Fatal("metadata custody", n)
						}
					}
				}
				if c.ErrorCode == "" {
					mask := func(source []byte) []byte {
						doc, er := losslessxml.Parse(source)
						if er != nil {
							t.Fatal(er)
						}
						for _, node := range doc.Elements() {
							if node.Name() != name(packaging.NSPresentationML, "pic") {
								continue
							}
							id, _, _, er := graphicsIdentity(doc, node, "nvPicPr")
							if er != nil {
								t.Fatal(er)
							}
							if id != c.Request.ShapeID {
								continue
							}
							var selected losslessxml.Element
							if family == "crop" {
								fill, er := graphicsOne(doc, node, packaging.NSPresentationML, "blipFill", false)
								if er != nil {
									t.Fatal(er)
								}
								selected, er = graphicsOne(doc, fill, packaging.NSDrawingML, "srcRect", true)
								if er != nil {
									t.Fatal(er)
								}
								if selected.Name().Local == "" {
									return source
								}
							} else {
								props, er := graphicsOne(doc, node, packaging.NSPresentationML, "spPr", false)
								if er != nil {
									t.Fatal(er)
								}
								selected, er = graphicsOne(doc, props, packaging.NSDrawingML, "xfrm", false)
								if er != nil {
									t.Fatal(er)
								}
							}
							a, z := selected.SourceRange()
							if family == "transform" {
								z, _ = selected.ContentRange()
							}
							return append(append([]byte{}, source[:a]...), source[z:]...)
						}
						t.Fatal("selected metadata target missing")
						return nil
					}
					if !bytes.Equal(mask(before[r.SlidePart]), mask(after[r.SlidePart])) {
						t.Fatal("XML changed outside selected metadata span")
					}
					reopened, er := OpenEditing(saved, packaging.Limits{})
					if er != nil {
						t.Fatal(er)
					}
					old, er := s.InspectPictures(r.SlidePart)
					if er != nil {
						t.Fatal(er)
					}
					got, er := reopened.InspectPictures(r.SlidePart)
					if er != nil || !reflect.DeepEqual(old, got) {
						t.Fatal("metadata reopened", er)
					}
					graphicsWriteOutput(t, root, "picture-"+family+"-"+c.CaseID+".pptx", saved)
				}
			})
		}
	}
}
