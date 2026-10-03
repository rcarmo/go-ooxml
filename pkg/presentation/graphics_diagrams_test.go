package presentation

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsDiagramRecipes(t *testing.T) {
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
				Nodes   []DiagramNode  `json:"nodes"`
				Edges   []DiagramEdge  `json:"edges"`
				Options DiagramOptions `json:"options"`
			} `json:"request"`
			Expected  *DiagramReceipt `json:"expected"`
			ErrorCode string          `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-diagrams.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 14 {
		t.Fatal("diagram inventory")
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
			receipt, e := s.AddDiagram(r.SlidePart, c.Request.Nodes, c.Request.Edges, c.Request.Options)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("diagram refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("diagram refusal generation")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if !reflect.DeepEqual(receipt, *c.Expected) || s.generation != version+1 {
					t.Fatalf("diagram receipt %#v want %#v", receipt, c.Expected)
				}
			}
			path := filepath.Join(t.TempDir(), "diagram.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			after := graphicsMembers(t, saved)
			if len(after) != len(before) {
				t.Fatal("diagram membership")
			}
			for n, b := range before {
				if c.ErrorCode != "" || n != r.SlidePart {
					if !bytes.Equal(b, after[n]) {
						t.Fatal("diagram custody", n)
					}
				}
			}
			if c.ErrorCode == "" {
				doc, er := losslessxml.Parse(after[r.SlidePart])
				if er != nil {
					t.Fatal(er)
				}
				selected := map[uint32]bool{}
				for _, n := range receipt.Nodes {
					selected[n.ShapeID] = true
				}
				for _, edge := range receipt.Edges {
					selected[edge.ShapeID] = true
				}
				spans := []losslessxml.Element{}
				for _, n := range doc.Elements() {
					nv := ""
					if n.Name() == name(packaging.NSPresentationML, "sp") {
						nv = "nvSpPr"
					}
					if n.Name() == name(packaging.NSPresentationML, "cxnSp") {
						nv = "nvCxnSpPr"
					}
					if nv == "" {
						continue
					}
					meta, er := graphicsOne(doc, n, packaging.NSPresentationML, nv, false)
					if er != nil {
						t.Fatal(er)
					}
					identity, er := graphicsOne(doc, meta, packaging.NSPresentationML, "cNvPr", false)
					if er != nil {
						t.Fatal(er)
					}
					raw, _ := graphicsAttr(identity, "", "id")
					id, _ := strconv.ParseUint(raw, 10, 32)
					if !selected[uint32(id)] {
						continue
					}
					spans = append(spans, n)
					if nv == "nvSpPr" {
						for i, node := range receipt.Nodes {
							if node.ShapeID != uint32(id) {
								continue
							}
							texts := []string{}
							for _, leaf := range doc.Elements() {
								if leaf.Name() == name(packaging.NSDrawingML, "t") && manipulationWithin(leaf, n) {
									v, _ := leaf.Text()
									texts = append(texts, v)
								}
							}
							want := strings.ReplaceAll(strings.ReplaceAll(c.Request.Nodes[i].Text, "\r\n", "\n"), "\r", "\n")
							if strings.Join(texts, "\n") != want {
								t.Fatal("diagram label")
							}
						}
					}
				}
				if len(spans) != len(selected) {
					t.Fatal("diagram appended spans")
				}
				restored := bytes.Clone(after[r.SlidePart])
				for i := len(spans) - 1; i >= 0; i-- {
					a, z := spans[i].SourceRange()
					restored = append(append([]byte{}, restored[:a]...), restored[z:]...)
				}
				if !bytes.Equal(restored, before[r.SlidePart]) {
					t.Fatal("diagram changed original XML")
				}
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				b, _, er := reopened.pkg.Part(r.SlidePart)
				if er != nil || !bytes.Equal(b, after[r.SlidePart]) {
					t.Fatal("diagram saved readback", er)
				}
				graphicsWriteOutput(t, root, "diagrams-"+c.CaseID+".pptx", saved)
				first := receipt.Nodes[0]
				textTarget, er := s.FindShape(r.SlidePart, first.ShapeID)
				if er != nil {
					t.Fatal("generated node native text target", er)
				}
				if er = s.SetShapeText(textTarget, "Editable label", false); er != nil {
					t.Fatal("generated node text edit", er)
				}
				formatTarget, er := s.FindFormatShape(r.SlidePart, first.ShapeID)
				if er != nil {
					t.Fatal("generated node native format target", er)
				}
				white := "FFFFFF"
				if er = s.SetRetainedShapeStyle(formatTarget, RetainedShapeStylePatch{Fill: &white}); er != nil {
					t.Fatal("generated node style edit", er)
				}
			}
		})
	}
}

func TestGraphicsPracticalDiagramQualitySample(t *testing.T) {
	root, assets := graphicsRoot(t)
	var recipe struct {
		BaseFixtureID string `json:"baseFixtureId"`
		SlidePart     string `json:"slidePart"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-diagrams.json"), &recipe); e != nil {
		t.Fatal(e)
	}
	input := graphicsReadAsset(t, root, assets, recipe.BaseFixtureID)
	s, e := OpenEditing(input, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	receipt, e := s.AddDiagram(recipe.SlidePart, []DiagramNode{{"input", "Input"}, {"process", "Process"}, {"output", "Output"}}, []DiagramEdge{{"input", "process"}, {"process", "output"}}, DiagramOptions{X: 457200, Y: 1828800, NodeWidth: 1828800, NodeHeight: 914400, Gap: 457200, Direction: "row"})
	if e != nil {
		t.Fatal(e)
	}
	if len(receipt.Nodes) != 3 || len(receipt.Edges) != 2 {
		t.Fatal("practical graph counts")
	}
	path := filepath.Join(t.TempDir(), "practical.pptx")
	if _, e = s.SaveAs(path); e != nil {
		t.Fatal(e)
	}
	saved, e := os.ReadFile(path)
	if e != nil {
		t.Fatal(e)
	}
	graphicsWriteOutput(t, root, "diagrams-practical.pptx", saved)
}
