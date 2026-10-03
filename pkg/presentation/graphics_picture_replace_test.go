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
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsPictureReplacementRecipes(t *testing.T) {
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
				ShapeID uint32 `json:"shapeId"`
				Payload struct {
					FixtureID  string `json:"fixtureId"`
					Part       string `json:"part"`
					ByteLength int    `json:"byteLength"`
					SHA256     string `json:"sha256"`
				} `json:"payload"`
				Options PictureReplacementOptions `json:"options"`
			} `json:"request"`
			RequestPatch struct {
				ShapeID uint32                     `json:"shapeId"`
				Options *PictureReplacementOptions `json:"options"`
			} `json:"requestPatch"`
			Expected  *PictureReplacementReceipt `json:"expected"`
			ErrorCode string                     `json:"errorCode"`
		} `json:"cases"`
	}
	if err := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-replacement.json"), &r); err != nil {
		t.Fatal(err)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 14 {
		t.Fatal("replacement inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			input := graphicsInput(t, base, c.Operations)
			session, err := OpenEditing(input, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			before := graphicsMembers(t, input)
			version := session.generation
			sourcePayload := graphicsMembers(t, graphicsReadAsset(t, root, assets, c.Request.Payload.FixtureID))[c.Request.Payload.Part]
			h := sha256.Sum256(sourcePayload)
			if len(sourcePayload) != c.Request.Payload.ByteLength || hex.EncodeToString(h[:]) != c.Request.Payload.SHA256 {
				t.Fatal("payload provenance")
			}
			payload := bytes.Clone(sourcePayload)
			id := c.Request.ShapeID
			options := c.Request.Options
			if c.RequestPatch.ShapeID != 0 {
				id = c.RequestPatch.ShapeID
			}
			if c.RequestPatch.Options != nil {
				options = *c.RequestPatch.Options
			}
			receipt, err := session.ReplacePicture(r.SlidePart, id, payload, options)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(err, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("refusal %v, want %s", err, c.ErrorCode)
				}
				if session.generation != version {
					t.Fatal("refusal changed generation")
				}
			} else {
				if err != nil {
					t.Fatal(err)
				}
				if receipt != *c.Expected {
					t.Fatalf("receipt %#v want %#v", receipt, c.Expected)
				}
				if session.generation != version+1 {
					t.Fatal("replacement version")
				}
				for i := range payload {
					payload[i] = 0
				}
			}
			path := filepath.Join(t.TempDir(), "replacement.pptx")
			if _, err = session.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			after := graphicsMembers(t, saved)
			added := 0
			if c.ErrorCode == "" {
				added = 1
			}
			if len(after) != len(before)+added {
				t.Fatal("membership")
			}
			for name, b := range before {
				allowed := c.ErrorCode == "" && (name == r.SlidePart || name == packaging.RelationshipsPathForPart(r.SlidePart) || name == packaging.ContentTypesPath)
				if !allowed && !bytes.Equal(after[name], b) {
					t.Fatal("unrelated payload", name)
				}
			}
			if c.ErrorCode == "" {
				if !bytes.Equal(after[receipt.MediaPart], sourcePayload) {
					t.Fatal("defensive media copy")
				}
				doc, err := losslessxml.Parse(after[r.SlidePart])
				if err != nil {
					t.Fatal(err)
				}
				var selected losslessxml.Element
				for _, node := range doc.Elements() {
					if node.Name() != name(packaging.NSPresentationML, "pic") {
						continue
					}
					got, _, _, e := graphicsIdentity(doc, node, "nvPicPr")
					if e != nil {
						t.Fatal(e)
					}
					if got == id {
						fill, e := graphicsOne(doc, node, packaging.NSPresentationML, "blipFill", false)
						if e != nil {
							t.Fatal(e)
						}
						selected, e = graphicsOne(doc, fill, packaging.NSDrawingML, "blip", false)
						if e != nil {
							t.Fatal(e)
						}
					}
				}
				restored, err := doc.Edit(nil, []losslessxml.AttributeEdit{{Target: selected, Name: name(packaging.NSDocumentRelationships, "embed"), Value: receipt.PreviousRelationshipID}})
				if err != nil || !bytes.Equal(restored, before[r.SlidePart]) {
					t.Fatal("replacement changed XML outside embed value", err)
				}
				oldSession, e := OpenEditing(input, packaging.Limits{})
				if e != nil {
					t.Fatal(e)
				}
				oldGraph, e := oldSession.pkg.Graph()
				if e != nil {
					t.Fatal(e)
				}
				nextGraph, e := session.pkg.Graph()
				if e != nil {
					t.Fatal(e)
				}
				for _, edge := range oldGraph.Edges {
					found := false
					for _, next := range nextGraph.Edges {
						if next == edge {
							found = true
							break
						}
					}
					if !found {
						t.Fatal("prior relationship changed")
					}
				}
				reopened, e := OpenEditing(saved, packaging.Limits{})
				if e != nil {
					t.Fatal(e)
				}
				a, e := session.InspectPictures(r.SlidePart)
				if e != nil {
					t.Fatal(e)
				}
				b, e := reopened.InspectPictures(r.SlidePart)
				if e != nil || !reflect.DeepEqual(a, b) {
					t.Fatal("saved replacement readback", e)
				}
				graphicsWriteOutput(t, root, "picture-replacement-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
