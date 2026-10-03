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

func TestGraphicsConnectorRecipes(t *testing.T) {
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
				Start   ConnectorEndpoint `json:"start"`
				End     ConnectorEndpoint `json:"end"`
				Options json.RawMessage   `json:"options"`
			} `json:"request"`
			Expected  *ConnectorReceipt `json:"expected"`
			ErrorCode string            `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-connectors.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 15 {
		t.Fatal("connector inventory")
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
			var options ConnectorOptions
			d := json.NewDecoder(bytes.NewReader(c.Request.Options))
			d.DisallowUnknownFields()
			e = d.Decode(&options)
			var receipt ConnectorReceipt
			if e != nil {
				e = editRefusal("PPTX_CONNECTOR_UNSUPPORTED", "native option decode")
			} else {
				receipt, e = s.AddConnector(r.SlidePart, c.Request.Start, c.Request.End, options)
			}
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("connector refusal %v want %s", e, c.ErrorCode)
				}
				if s.generation != version {
					t.Fatal("refusal version")
				}
			} else {
				if e != nil {
					t.Fatal(e)
				}
				if !reflect.DeepEqual(receipt, *c.Expected) || s.generation != version+1 {
					t.Fatalf("connector receipt %#v want %#v", receipt, c.Expected)
				}
			}
			path := filepath.Join(t.TempDir(), "connector.pptx")
			if _, e = s.SaveAs(path); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(path)
			if e != nil {
				t.Fatal(e)
			}
			after := graphicsMembers(t, saved)
			if len(after) != len(before) {
				t.Fatal("connector membership")
			}
			for n, b := range before {
				if c.ErrorCode != "" || n != r.SlidePart {
					if !bytes.Equal(b, after[n]) {
						t.Fatal("connector custody", n)
					}
				}
			}
			if c.ErrorCode == "" {
				doc, er := losslessxml.Parse(after[r.SlidePart])
				if er != nil {
					t.Fatal(er)
				}
				var connector losslessxml.Element
				for _, n := range doc.Elements() {
					if n.Name() != name(packaging.NSPresentationML, "cxnSp") {
						continue
					}
					nv, er := graphicsOne(doc, n, packaging.NSPresentationML, "nvCxnSpPr", false)
					if er != nil {
						t.Fatal(er)
					}
					id, er := graphicsOne(doc, nv, packaging.NSPresentationML, "cNvPr", false)
					if er != nil {
						t.Fatal(er)
					}
					raw, _ := graphicsAttr(id, "", "id")
					v, _ := graphicsNumber(raw, 1, 2147483647)
					if uint32(v) == receipt.ShapeID {
						connector = n
					}
				}
				a, z := connector.SourceRange()
				restored := append(append([]byte{}, after[r.SlidePart][:a]...), after[r.SlidePart][z:]...)
				if !bytes.Equal(restored, before[r.SlidePart]) {
					t.Fatal("connector changed prior XML")
				}
				nv, er := graphicsOne(doc, connector, packaging.NSPresentationML, "nvCxnSpPr", false)
				if er != nil {
					t.Fatal(er)
				}
				pr, er := graphicsOne(doc, nv, packaging.NSPresentationML, "cNvCxnSpPr", false)
				if er != nil {
					t.Fatal(er)
				}
				for _, v := range []struct {
					local    string
					endpoint ConnectorPoint
				}{{"stCxn", receipt.Start}, {"endCxn", receipt.End}} {
					n, er := graphicsOne(doc, pr, packaging.NSDrawingML, v.local, false)
					if er != nil {
						t.Fatal(er)
					}
					id, _ := graphicsAttr(n, "", "id")
					site, _ := graphicsAttr(n, "", "idx")
					wantID, _ := json.Marshal(v.endpoint.ShapeID)
					wantSite, _ := json.Marshal(v.endpoint.Site)
					if id != string(wantID) || site != string(wantSite) {
						t.Fatal("connector attachment")
					}
				}
				reopened, er := OpenEditing(saved, packaging.Limits{})
				if er != nil {
					t.Fatal(er)
				}
				b, _, er := reopened.pkg.Part(r.SlidePart)
				if er != nil || !bytes.Equal(b, after[r.SlidePart]) {
					t.Fatal("saved connector XML", er)
				}
				graphicsWriteOutput(t, root, "connectors-"+c.CaseID+".pptx", saved)
			}
		})
	}
}
