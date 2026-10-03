package presentation

import (
	"encoding/json"
	"errors"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsSmartArtSourceRefusalControls(t *testing.T) {
	root, assets := graphicsRoot(t)
	var r struct {
		FixtureID string `json:"fixtureId"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-smartart-office-source.json"), &r); e != nil {
		t.Fatal(e)
	}
	input := graphicsReadAsset(t, root, assets, r.FixtureID)
	for _, variant := range []string{"duplicate nonzero drawing IDs", "ambiguous owner", "dangling model reference"} {
		t.Run(variant, func(t *testing.T) {
			members := graphicsMembers(t, input)
			switch variant {
			case "duplicate nonzero drawing IDs":
				members["ppt/diagrams/drawing1.xml"] = []byte(strings.ReplaceAll(string(members["ppt/diagrams/drawing1.xml"]), `id="0"`, `id="5"`))
			case "ambiguous owner":
				members["ppt/diagrams/_rels/data1.xml.rels"] = []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>`)
			case "dangling model reference":
				members["ppt/diagrams/drawing1.xml"] = []byte(strings.Replace(string(members["ppt/diagrams/drawing1.xml"]), `modelId="`, `modelId="{00000000-0000-0000-0000-000000000000}" fake="`, 1))
			}
			source, e := OpenEditing(graphicsArchive(t, members), packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target := graphicsCopyTarget(t, "Refusal")
			before := graphicsSessionMembers(t, target)
			version := target.generation
			_, e = target.CopySmartArtFrom("ppt/slides/slide1.xml", source, "ppt/slides/slide1.xml", 4)
			var refusal *packaging.Refusal
			if !errors.As(e, &refusal) || refusal.Kind != "PPTX_SMARTART_UNSUPPORTED" {
				t.Fatalf("source refusal %v", e)
			}
			if target.generation != version || !reflect.DeepEqual(before, graphicsSessionMembers(t, target)) {
				t.Fatal("source refusal atomicity")
			}
		})
	}
}
