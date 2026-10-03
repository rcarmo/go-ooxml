package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"
)

func TestGraphicsAutoShapeRecipes(t *testing.T) {
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
				Preset   string          `json:"preset"`
				Geometry PictureGeometry `json:"geometry"`
				Options  json.RawMessage `json:"options"`
			} `json:"request"`
			Expected  *AutoShapeReceipt `json:"expected"`
			ErrorCode string            `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-autoshapes.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 18 {
		t.Fatal("preset inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			s, e := OpenEditing(graphicsInput(t, base, c.Operations), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsSessionMembers(t, s)
			version := s.generation
			var o AutoShapeOptions
			d := json.NewDecoder(bytes.NewReader(c.Request.Options))
			d.DisallowUnknownFields()
			e = d.Decode(&o)
			var receipt AutoShapeReceipt
			if e != nil {
				e = editRefusal("PPTX_AUTOSHAPE_UNSUPPORTED", "native option decode")
			} else {
				receipt, e = s.AddAutoShape(r.SlidePart, c.Request.Preset, c.Request.Geometry, o)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("preset refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, s)) {
					t.Fatal("preset refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if !reflect.DeepEqual(receipt, *c.Expected) || s.generation != version+1 {
				t.Fatal("preset receipt", receipt, c.Expected)
			}
			after := graphicsSessionMembers(t, s)
			doc, shape := graphicsNewShapeCustody(t, before, after, r.SlidePart, receipt.ShapeID)
			props, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
			if e != nil {
				t.Fatal(e)
			}
			preset, e := graphicsOne(doc, props, packaging.NSDrawingML, "prstGeom", false)
			if e != nil {
				t.Fatal(e)
			}
			kind, _ := graphicsAttr(preset, "", "prst")
			if kind != c.Request.Preset {
				t.Fatal("preset kind")
			}
			av, e := graphicsOne(doc, preset, packaging.NSDrawingML, "avLst", false)
			if e != nil {
				t.Fatal(e)
			}
			if len(graphicsChildren(doc, av)) != len(receipt.Adjustments) {
				t.Fatal("preset adjustment count")
			}
			for _, guide := range graphicsChildren(doc, av) {
				k, _ := graphicsAttr(guide, "", "name")
				value, _ := graphicsAttr(guide, "", "fmla")
				if value != "val "+strconv.FormatInt(receipt.Adjustments[k], 10) {
					t.Fatal("preset guide")
				}
			}
			text := []string{}
			for _, n := range doc.Elements() {
				if n.Name() == name(packaging.NSDrawingML, "t") && manipulationWithin(n, shape) {
					v, _ := n.Text()
					text = append(text, v)
				}
			}
			want := ""
			if o.Text != nil {
				want = strings.ReplaceAll(strings.ReplaceAll(*o.Text, "\r\n", "\n"), "\r", "\n")
			}
			if strings.Join(text, "\n") != want {
				t.Fatal("preset text")
			}
			path := filepath.Join(t.TempDir(), "shape.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			reopened, e := OpenEditing(saved, packaging.Limits{})
			if e != nil || !reflect.DeepEqual(after, graphicsSessionMembers(t, reopened)) {
				t.Fatal("preset reopen", e)
			}
			graphicsWriteOutput(t, root, "autoshapes-"+c.CaseID+".pptx", saved)
		})
	}
}
func graphicsNewShapeCustody(t *testing.T, before, after map[string][]byte, part string, id uint32) (*losslessxml.Document, losslessxml.Element) {
	t.Helper()
	if len(before) != len(after) {
		t.Fatal("shape membership")
	}
	for n, b := range before {
		if n != part && !bytes.Equal(b, after[n]) {
			t.Fatal("shape unrelated bytes", n)
		}
	}
	doc, e := losslessxml.Parse(after[part])
	if e != nil {
		t.Fatal(e)
	}
	var shape losslessxml.Element
	for _, n := range doc.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "sp") {
			continue
		}
		got, _, _, e := graphicsIdentity(doc, n, "nvSpPr")
		if e != nil {
			t.Fatal(e)
		}
		if got == id {
			shape = n
		}
	}
	a, z := shape.SourceRange()
	if a < 0 {
		t.Fatal("authored shape missing")
	}
	restored := append(append([]byte{}, after[part][:a]...), after[part][z:]...)
	if !bytes.Equal(restored, before[part]) {
		t.Fatal("shape changed old XML")
	}
	return doc, shape
}
