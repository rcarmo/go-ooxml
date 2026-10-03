package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"reflect"
	"testing"
)

func TestGraphicsFreeformRecipes(t *testing.T) {
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
				Geometry PictureGeometry `json:"geometry"`
				Path     json.RawMessage `json:"path"`
				Options  json.RawMessage `json:"options"`
			} `json:"request"`
			Expected  *FreeformReceipt `json:"expected"`
			ErrorCode string           `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-freeform.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 15 {
		t.Fatal("freeform inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			s, e := OpenEditing(graphicsInput(t, base, c.Operations), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			before := graphicsSessionMembers(t, s)
			version := s.generation
			var path FreeformPath
			var o FreeformOptions
			d := json.NewDecoder(bytes.NewReader(c.Request.Path))
			d.DisallowUnknownFields()
			e = d.Decode(&path)
			if e == nil {
				d = json.NewDecoder(bytes.NewReader(c.Request.Options))
				d.DisallowUnknownFields()
				e = d.Decode(&o)
			}
			var receipt FreeformReceipt
			if e != nil {
				e = editRefusal("PPTX_FREEFORM_UNSUPPORTED", "native path/options decode")
			} else {
				receipt, e = s.AddFreeform(r.SlidePart, c.Request.Geometry, path, o)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("freeform refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, s)) {
					t.Fatal("freeform refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if !reflect.DeepEqual(receipt, *c.Expected) || s.generation != version+1 {
				t.Fatalf("freeform receipt %#v want %#v", receipt, c.Expected)
			}
			after := graphicsSessionMembers(t, s)
			doc, shape := graphicsNewShapeCustody(t, before, after, r.SlidePart, receipt.ShapeID)
			props, e := graphicsOne(doc, shape, packaging.NSPresentationML, "spPr", false)
			if e != nil {
				t.Fatal(e)
			}
			custom, e := graphicsOne(doc, props, packaging.NSDrawingML, "custGeom", false)
			if e != nil {
				t.Fatal(e)
			}
			paths, e := graphicsOne(doc, custom, packaging.NSDrawingML, "pathLst", false)
			if e != nil {
				t.Fatal(e)
			}
			node, e := graphicsOne(doc, paths, packaging.NSDrawingML, "path", false)
			if e != nil {
				t.Fatal(e)
			}
			commands := graphicsChildren(doc, node)
			if len(commands) != len(path.Commands) {
				t.Fatal("freeform command count")
			}
			for i, n := range commands {
				want := map[string]string{"move": "moveTo", "line": "lnTo", "close": "close"}[path.Commands[i].Op]
				if n.Name() != name(packaging.NSDrawingML, want) {
					t.Fatal("freeform command order")
				}
			}
			var out bytes.Buffer
			if e = s.pkg.WriteTo(&out); e != nil {
				t.Fatal(e)
			}
			reopened, e := OpenEditing(out.Bytes(), packaging.Limits{})
			if e != nil || !reflect.DeepEqual(after, graphicsSessionMembers(t, reopened)) {
				t.Fatal("freeform reopen", e)
			}
			graphicsWriteOutput(t, root, "freeform-"+c.CaseID+".pptx", out.Bytes())
		})
	}
}
