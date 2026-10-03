package presentation

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type graphicsInsertCase struct {
	ScenarioID string              `json:"scenarioId"`
	CaseID     string              `json:"caseId"`
	Operations []graphicsOperation `json:"operations"`
	Request    struct {
		Payload struct {
			FixtureID  string `json:"fixtureId"`
			Part       string `json:"part"`
			ByteLength int    `json:"byteLength"`
			SHA256     string `json:"sha256"`
		} `json:"payload"`
		Geometry map[string]json.RawMessage `json:"geometry"`
		Options  map[string]json.RawMessage `json:"options"`
	} `json:"request"`
	RequestPatch struct {
		Geometry      map[string]json.RawMessage `json:"geometry"`
		Options       map[string]json.RawMessage `json:"options"`
		PayloadChange string                     `json:"payloadChange"`
	} `json:"requestPatch"`
	Expected []struct {
		ShapeID        uint32          `json:"shapeId"`
		SlidePart      string          `json:"slidePart"`
		MediaPart      string          `json:"mediaPart"`
		RelationshipID string          `json:"relationshipId"`
		Target         string          `json:"target"`
		ContentType    string          `json:"contentType"`
		Name           string          `json:"name"`
		Description    string          `json:"description"`
		Geometry       PictureGeometry `json:"geometry"`
	} `json:"expected"`
	ErrorCode string `json:"errorCode"`
}

func TestGraphicsPictureInsertionRecipes(t *testing.T) {
	root, assets := graphicsRoot(t)
	var recipe struct {
		Feature       string               `json:"feature"`
		Contract      string               `json:"contract"`
		BaseFixtureID string               `json:"baseFixtureId"`
		SlidePart     string               `json:"slidePart"`
		Cases         []graphicsInsertCase `json:"cases"`
	}
	if err := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-insertion.json"), &recipe); err != nil {
		t.Fatal(err)
	}
	graphicsReadAsset(t, root, assets, recipe.Feature)
	graphicsReadAsset(t, root, assets, recipe.Contract)
	base := graphicsReadAsset(t, root, assets, recipe.BaseFixtureID)
	if len(recipe.Cases) != 18 {
		t.Fatal("insertion inventory")
	}
	for _, c := range recipe.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			input := graphicsInput(t, base, c.Operations)
			session, err := OpenEditing(input, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			before := graphicsMembers(t, input)
			generation := session.generation
			payloadParts := graphicsMembers(t, graphicsReadAsset(t, root, assets, c.Request.Payload.FixtureID))
			originalPayload := payloadParts[c.Request.Payload.Part]
			payloadHash := sha256.Sum256(originalPayload)
			if len(originalPayload) != c.Request.Payload.ByteLength || hex.EncodeToString(payloadHash[:]) != c.Request.Payload.SHA256 {
				t.Fatal("payload provenance length/hash")
			}
			payload := bytes.Clone(originalPayload)
			if c.RequestPatch.PayloadChange == "empty" {
				payload = nil
			}
			for k, v := range c.RequestPatch.Geometry {
				c.Request.Geometry[k] = v
			}
			for k, v := range c.RequestPatch.Options {
				c.Request.Options[k] = v
			}
			geometryJSON, _ := json.Marshal(c.Request.Geometry)
			optionsJSON, _ := json.Marshal(c.Request.Options)
			var geometry PictureGeometry
			var options PictureOptions
			decodeErr := json.Unmarshal(geometryJSON, &geometry)
			decoder := json.NewDecoder(bytes.NewReader(optionsJSON))
			decoder.DisallowUnknownFields()
			optionErr := decoder.Decode(&options)
			var receipt PictureReceipt
			if decodeErr != nil || optionErr != nil {
				err = editRefusal("PPTX_PICTURE_UNSUPPORTED", "native request decode refused")
			} else {
				receipt, err = session.AddPicture(recipe.SlidePart, payload, geometry, options)
			}
			if c.ErrorCode != "" {
				var r *packaging.Refusal
				if !errors.As(err, &r) || r.Kind != c.ErrorCode {
					t.Fatalf("refusal %v, want %s", err, c.ErrorCode)
				}
				if session.generation != generation {
					t.Fatal("refusal changed version")
				}
			} else {
				if err != nil {
					t.Fatal(err)
				}
				for i, want := range c.Expected {
					if i > 0 {
						receipt, err = session.AddPicture(recipe.SlidePart, bytes.Clone(originalPayload), geometry, options)
						if err != nil {
							t.Fatal(err)
						}
					}
					if receipt != (PictureReceipt{want.ShapeID, want.SlidePart, want.MediaPart, want.RelationshipID}) {
						t.Fatalf("receipt %#v differs %#v", receipt, want)
					}
					for j := range payload {
						payload[j] = 0
					}
					info, err := session.InspectPictures(recipe.SlidePart)
					if err != nil {
						t.Fatal(err)
					}
					var selected *PictureInfo
					for j := range info {
						if info[j].ShapeID == want.ShapeID {
							selected = &info[j]
						}
					}
					if selected == nil || selected.Name != want.Name || selected.Description == nil || *selected.Description != want.Description || selected.Embedded == nil || selected.Embedded.PartName != want.MediaPart || selected.Embedded.RelationshipID != want.RelationshipID || selected.Embedded.Target != want.Target || selected.Embedded.ContentType != want.ContentType || selected.Transform == nil || selected.Transform.X != want.Geometry.X || selected.Transform.Y != want.Geometry.Y || selected.Transform.Width != want.Geometry.Width || selected.Transform.Height != want.Geometry.Height {
						t.Fatalf("inserted literal metadata mismatch %#v", selected)
					}
				}
			}
			path := filepath.Join(t.TempDir(), "pictures.pptx")
			if _, err = session.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			after := graphicsMembers(t, saved)
			for name, b := range before {
				allowed := c.ErrorCode == "" && (name == recipe.SlidePart || name == packaging.RelationshipsPathForPart(recipe.SlidePart) || name == packaging.ContentTypesPath)
				if !allowed && !bytes.Equal(after[name], b) {
					t.Fatal("unrelated payload changed", name)
				}
			}
			expectedAdded := 0
			if c.ErrorCode == "" {
				expectedAdded = len(c.Expected)
			}
			if len(after) != len(before)+expectedAdded {
				t.Fatal("unexpected package membership")
			}
			for _, want := range c.Expected {
				if !bytes.Equal(after[want.MediaPart], originalPayload) {
					t.Fatal("media changed or caller aliased", want.MediaPart)
				}
			}
			if c.ErrorCode == "" {
				// Remove exactly the appended nodes and compare all existing slide XML.
				changed := after[recipe.SlidePart]
				doc, err := losslessxml.Parse(changed)
				if err != nil {
					t.Fatal(err)
				}
				var remove []losslessxml.Element
				for _, node := range doc.Elements() {
					if node.Name() != name(packaging.NSPresentationML, "pic") {
						continue
					}
					id, _, _, e := graphicsIdentity(doc, node, "nvPicPr")
					if e != nil {
						t.Fatal(e)
					}
					for _, want := range c.Expected {
						if id == want.ShapeID {
							remove = append(remove, node)
						}
					}
				}
				if len(remove) != len(c.Expected) {
					t.Fatal("appended picture count")
				}
				for i := len(remove) - 1; i >= 0; i-- {
					a, b := remove[i].SourceRange()
					changed = append(append([]byte{}, changed[:a]...), changed[b:]...)
				}
				if !bytes.Equal(changed, before[recipe.SlidePart]) {
					t.Fatal("existing slide XML rewritten")
				}
				oldGraph, err := OpenEditing(input, packaging.Limits{})
				if err != nil {
					t.Fatal(err)
				}
				oldEdges, err := oldGraph.pkg.Graph()
				if err != nil {
					t.Fatal(err)
				}
				nextEdges, err := session.pkg.Graph()
				if err != nil {
					t.Fatal(err)
				}
				for _, edge := range oldEdges.Edges {
					found := false
					for _, got := range nextEdges.Edges {
						if got == edge {
							found = true
							break
						}
					}
					if !found {
						t.Fatal("existing relationship changed", edge)
					}
				}
				if c.CaseID == "conflicting-MIME" && !bytes.Contains(after[packaging.ContentTypesPath], []byte(`Extension="png" ContentType="application/octet-stream"`)) {
					t.Fatal("conflicting default was rewritten")
				}
				if output := os.Getenv("OOXML_GRAPHICS_OUTPUT"); output != "" {
					dir, e := filepath.Abs(output)
					if e != nil {
						t.Fatal(e)
					}
					rel, e := filepath.Rel(root, dir)
					if e != nil || rel == "." || rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
						t.Fatal("output inside shared root")
					}
					if e = os.MkdirAll(dir, 0755); e != nil {
						t.Fatal(e)
					}
					if e = os.WriteFile(filepath.Join(dir, "picture-insertion-"+c.CaseID+".pptx"), saved, 0600); e != nil {
						t.Fatal(e)
					}
				}
				reopened, err := OpenEditing(saved, packaging.Limits{})
				if err != nil {
					t.Fatal(err)
				}
				a, err := session.InspectPictures(recipe.SlidePart)
				if err != nil {
					t.Fatal(err)
				}
				b, err := reopened.InspectPictures(recipe.SlidePart)
				if err != nil || !reflect.DeepEqual(a, b) {
					t.Fatal("saved picture readback differs", err)
				}
			}
		})
	}
}
