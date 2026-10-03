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

func TestGraphicsPictureSVGRecipes(t *testing.T) {
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
				SVG      string `json:"svg"`
				Fallback struct {
					FixtureID  string `json:"fixtureId"`
					Part       string `json:"part"`
					ByteLength int    `json:"byteLength"`
					SHA256     string `json:"sha256"`
				} `json:"fallback"`
				Geometry PictureGeometry `json:"geometry"`
				Options  PictureOptions  `json:"options"`
			} `json:"request"`
			RequestPatch struct {
				EmptyFallback bool `json:"emptyFallback"`
			} `json:"requestPatch"`
			Expected  *SvgPictureReceipt `json:"expected"`
			ErrorCode string             `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-svg.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 11 {
		t.Fatal("SVG inventory")
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
			svg := []byte(c.Request.SVG)
			originalSVG := bytes.Clone(svg)
			fallback := graphicsMembers(t, graphicsReadAsset(t, root, assets, c.Request.Fallback.FixtureID))[c.Request.Fallback.Part]
			h := sha256.Sum256(fallback)
			if len(fallback) != c.Request.Fallback.ByteLength || hex.EncodeToString(h[:]) != c.Request.Fallback.SHA256 {
				t.Fatal("fallback provenance")
			}
			originalFallback := bytes.Clone(fallback)
			if c.RequestPatch.EmptyFallback {
				fallback = nil
			}
			receipt, e := s.AddSvgPicture(r.SlidePart, svg, fallback, c.Request.Geometry, c.Request.Options)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("SVG refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("SVG refusal version")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if receipt != *c.Expected || s.generation != version+1 {
					t.Fatal("SVG receipt/version", receipt, c.Expected)
				}
				for i := range svg {
					svg[i] = 0
				}
				for i := range fallback {
					fallback[i] = 0
				}
			}
			path := filepath.Join(t.TempDir(), "svg.pptx")
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
				extra = 2
			}
			if len(after) != len(before)+extra {
				t.Fatal("SVG membership")
			}
			for n, b := range before {
				allowed := c.ErrorCode == "" && (n == r.SlidePart || n == packaging.RelationshipsPathForPart(r.SlidePart) || n == packaging.ContentTypesPath)
				if !allowed && !bytes.Equal(after[n], b) {
					t.Fatal("SVG custody", n)
				}
			}
			if c.ErrorCode == "" {
				if !bytes.Equal(after[receipt.SvgMediaPart], originalSVG) || !bytes.Equal(after[receipt.MediaPart], originalFallback) {
					t.Fatal("SVG pair payload custody")
				}
				doc, er := losslessxml.Parse(after[r.SlidePart])
				if er != nil {
					t.Fatal(er)
				}
				paired := false
				for _, n := range doc.Elements() {
					if n.Name() == name("http://schemas.microsoft.com/office/drawing/2016/SVG/main", "svgBlip") {
						value, _ := graphicsAttr(n, packaging.NSDocumentRelationships, "embed")
						if value == receipt.SvgRelationshipID {
							paired = true
						}
					}
				}
				if !paired {
					t.Fatal("SVG relationship leaf")
				}
				g, er := s.pkg.Graph()
				if er != nil {
					t.Fatal(er)
				}
				count := 0
				for _, edge := range g.Edges {
					if edge.Source == r.SlidePart && edge.Type == packaging.RelTypeImage && (edge.ID == receipt.RelationshipID && edge.ResolvedPart == receipt.MediaPart || edge.ID == receipt.SvgRelationshipID && edge.ResolvedPart == receipt.SvgMediaPart) {
						count++
					}
				}
				if count != 2 {
					t.Fatal("SVG pair graph")
				}
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				a, er := s.InspectPictures(r.SlidePart)
				if er != nil {
					t.Fatal(er)
				}
				b, er := reopened.InspectPictures(r.SlidePart)
				if er != nil || !reflect.DeepEqual(a, b) {
					t.Fatal("SVG saved readback", er)
				}
				graphicsWriteOutput(t, root, "picture-svg-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
