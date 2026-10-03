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

func TestGraphicsShapeGroupingRecipes(t *testing.T) {
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
				ShapeIDs []uint32        `json:"shapeIds"`
				Geometry PictureGeometry `json:"geometry"`
				Options  json.RawMessage `json:"options"`
			} `json:"request"`
			Expected  *ShapeGroupReceipt `json:"expected"`
			ErrorCode string             `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-shape-group.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 16 {
		t.Fatal("group inventory")
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
			var options ShapeGroupOptions
			d := json.NewDecoder(bytes.NewReader(c.Request.Options))
			d.DisallowUnknownFields()
			e = d.Decode(&options)
			var receipt ShapeGroupReceipt
			if e != nil {
				e = editRefusal("PPTX_GROUP_UNSUPPORTED", "native group option decode")
			} else {
				receipt, e = s.GroupShapes(r.SlidePart, c.Request.ShapeIDs, c.Request.Geometry, options)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("group refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("group refusal version")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if !reflect.DeepEqual(receipt, *c.Expected) || s.generation != version+1 {
					t.Fatalf("group receipt %#v want %#v", receipt, c.Expected)
				}
			}
			path := filepath.Join(t.TempDir(), "group.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			after := graphicsMembers(t, saved)
			if len(after) != len(before) {
				t.Fatal("group membership")
			}
			for n, b := range before {
				if c.ErrorCode != "" || n != r.SlidePart {
					if !bytes.Equal(after[n], b) {
						t.Fatal("group unrelated bytes", n)
					}
				}
			}
			if c.ErrorCode == "" {
				doc, er := losslessxml.Parse(after[r.SlidePart])
				if er != nil {
					t.Fatal(er)
				}
				var group losslessxml.Element
				for _, n := range doc.Elements() {
					if n.Name() != name(packaging.NSPresentationML, "grpSp") {
						continue
					}
					id, _, _, er := graphicsIdentity(doc, n, "nvGrpSpPr")
					if er != nil {
						t.Fatal(er)
					}
					if id == receipt.ShapeID {
						group = n
					}
				}
				children := graphicsChildren(doc, group)
				if len(children) != 4 {
					t.Fatal("group children")
				}
				a, _ := children[2].SourceRange()
				_, z := children[3].SourceRange()
				start, end := group.SourceRange()
				restored := append(append(append([]byte{}, after[r.SlidePart][:start]...), after[r.SlidePart][a:z]...), after[r.SlidePart][end:]...)
				if !bytes.Equal(restored, before[r.SlidePart]) {
					t.Fatal("group child slice/outer XML changed")
				}
				tr, er := graphicsReadTransform(doc, children[1], true)
				if er != nil || tr == nil || tr.X != receipt.Geometry.X || tr.ChildX != receipt.Geometry.X || tr.Width != receipt.Geometry.Width || tr.ChildWidth != receipt.Geometry.Width || tr.Y != receipt.Geometry.Y || tr.ChildY != receipt.Geometry.Y || tr.Height != receipt.Geometry.Height || tr.ChildHeight != receipt.Geometry.Height {
					t.Fatal("identity group transform", tr, er)
				}
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				pics, er := s.InspectPictures(r.SlidePart)
				if er != nil {
					t.Fatal(er)
				}
				read, er := reopened.InspectPictures(r.SlidePart)
				if er != nil || !reflect.DeepEqual(pics, read) {
					t.Fatal("group picture ancestry readback", er)
				}
				graphicsWriteOutput(t, root, "shape-group-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
