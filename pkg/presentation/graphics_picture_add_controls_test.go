package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"math"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsPictureInsertionNativeControls(t *testing.T) {
	root, assets := graphicsRoot(t)
	var recipe struct {
		BaseFixtureID string               `json:"baseFixtureId"`
		SlidePart     string               `json:"slidePart"`
		Cases         []graphicsInsertCase `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-picture-insertion.json"), &recipe); e != nil {
		t.Fatal(e)
	}
	base := graphicsReadAsset(t, root, assets, recipe.BaseFixtureID)
	parts := graphicsMembers(t, graphicsReadAsset(t, root, assets, recipe.Cases[0].Request.Payload.FixtureID))
	payload := parts[recipe.Cases[0].Request.Payload.Part]
	g := PictureGeometry{0, 0, 914400, 914400}
	o := PictureOptions{ContentType: "image/png"}
	for _, c := range []struct {
		name     string
		payload  []byte
		geometry PictureGeometry
		options  PictureOptions
	}{{"oversized", make([]byte, 64*1024*1024+1), g, o}, {"negative-width", payload, PictureGeometry{0, 0, -1, 1}, o}, {"max-width", payload, PictureGeometry{0, 0, math.MaxInt32 + 1, 1}, o}, {"invalid-description", payload, g, PictureOptions{ContentType: "image/png", Description: func() *string { s := "bad\x00"; return &s }()}}} {
		t.Run(c.name, func(t *testing.T) {
			s, e := OpenEditing(base, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			version := s.generation
			_, e = s.AddPicture(recipe.SlidePart, c.payload, c.geometry, c.options)
			var refusal *packaging.Refusal
			if !errors.As(e, &refusal) || refusal.Kind != "PPTX_PICTURE_UNSUPPORTED" {
				t.Fatalf("refusal %v", e)
			}
			if s.generation != version {
				t.Fatal("refusal changed generation")
			}
			output := filepath.Join(t.TempDir(), "unchanged.pptx")
			if _, e = s.SaveAs(output); e != nil {
				t.Fatal(e)
			}
			b, e := os.ReadFile(output)
			if e != nil || !reflect.DeepEqual(graphicsMembers(t, b), graphicsMembers(t, base)) {
				t.Fatal("native refusal changed package", e)
			}
			if _, e = s.AddPicture(recipe.SlidePart, payload, g, o); e != nil {
				t.Fatal("refusal invalidated session", e)
			}
		})
	}
	t.Run("default-metadata", func(t *testing.T) {
		s, e := OpenEditing(base, packaging.Limits{})
		if e != nil {
			t.Fatal(e)
		}
		receipt, e := s.AddPicture(recipe.SlidePart, payload, g, o)
		if e != nil {
			t.Fatal(e)
		}
		info, e := s.InspectPictures(recipe.SlidePart)
		if e != nil {
			t.Fatal(e)
		}
		for _, p := range info {
			if p.ShapeID == receipt.ShapeID {
				if p.Name != "Picture 1034" || p.Description != nil {
					t.Fatalf("default metadata %#v", p)
				}
				return
			}
		}
		t.Fatal("appended picture absent")
	})
	t.Run("late-graph-plan-refusal", func(t *testing.T) {
		signed := graphicsMembers(t, base)
		signed["_xmlsignatures/origin.sigs"] = []byte("opaque signature marker")
		signed[packaging.ContentTypesPath] = []byte(strings.Replace(string(signed[packaging.ContentTypesPath]), "</Types>", `<Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/></Types>`, 1))
		input := graphicsArchive(t, signed)
		s, e := OpenEditing(input, packaging.Limits{})
		if e != nil {
			t.Fatal(e)
		}
		version := s.generation
		_, e = s.AddPicture(recipe.SlidePart, payload, g, o)
		var refusal *packaging.Refusal
		if !errors.As(e, &refusal) || refusal.Kind != "unsupported_structure" {
			t.Fatalf("signed graph refusal %v", e)
		}
		if s.generation != version {
			t.Fatal("late refusal changed generation")
		}
		output := filepath.Join(t.TempDir(), "signed.pptx")
		if _, e = s.SaveAs(output); e != nil {
			t.Fatal(e)
		}
		b, e := os.ReadFile(output)
		if e != nil || !bytes.Equal(b, input) {
			t.Fatal("late plan refusal dirtied package", e)
		}
	})
}
