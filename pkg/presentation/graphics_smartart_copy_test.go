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

func graphicsCopyTarget(t *testing.T, title string) *EditSession {
	t.Helper()
	creator, e := NewContractPresentation()
	if e != nil {
		t.Fatal(e)
	}
	defer creator.Close()
	if e = creator.AddTitleSlide(title, ""); e != nil {
		t.Fatal(e)
	}
	path := filepath.Join(t.TempDir(), "target.pptx")
	if e = creator.SaveAs(path); e != nil {
		t.Fatal(e)
	}
	b, e := os.ReadFile(path)
	if e != nil {
		t.Fatal(e)
	}
	s, e := OpenEditing(b, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	// Shared destination recipe is a title-only slide with initial max ID 2.
	// Go's creator emits an extra empty subtitle. Remove that freshly authored
	// placeholder in memory; never normalise or edit a sealed source fixture.
	raw, h, e := s.pkg.Part("ppt/slides/slide1.xml")
	if e != nil {
		t.Fatal(e)
	}
	doc, e := losslessxml.Parse(raw)
	if e != nil {
		t.Fatal(e)
	}
	var empty []losslessxml.Element
	for _, n := range doc.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "sp") {
			continue
		}
		text := ""
		for _, leaf := range doc.Elements() {
			if leaf.Name() == name(packaging.NSDrawingML, "t") && manipulationWithin(leaf, n) {
				v, _ := leaf.Text()
				text += v
			}
		}
		if text == "" {
			empty = append(empty, n)
		}
	}
	if len(empty) != 1 {
		t.Fatal("title-only destination assembly")
	}
	raw, e = doc.RemoveElements(empty)
	if e != nil {
		t.Fatal(e)
	}
	if e = s.pkg.Replace([]packaging.Replacement{{Part: "ppt/slides/slide1.xml", ExpectedSHA256: h, Data: raw}}); e != nil {
		t.Fatal(e)
	}
	return s
}
func graphicsSessionMembers(t *testing.T, s *EditSession) map[string][]byte {
	t.Helper()
	var b bytes.Buffer
	if e := s.pkg.WriteTo(&b); e != nil {
		t.Fatal(e)
	}
	return graphicsMembers(t, b.Bytes())
}
func graphicsVerifySmartCopy(t *testing.T, source, target *EditSession, sourcePart, targetPart string, id uint32, receipt SmartArtCopyReceipt, sourceBefore, targetBefore map[string][]byte, same bool) {
	t.Helper()
	a, e := source.InspectSmartArt(sourcePart)
	if e != nil {
		t.Fatal(e)
	}
	var original SmartArtInfo
	for _, g := range a {
		if g.ShapeID == id {
			original = g
			break
		}
	}
	b, e := target.InspectSmartArt(targetPart)
	if e != nil {
		t.Fatal(e)
	}
	var copied *SmartArtInfo
	for i := range b {
		if b[i].ShapeID == receipt.ShapeID {
			copied = &b[i]
		}
	}
	if copied == nil || len(copied.Parts) != len(original.Parts) || len(copied.Roots) != len(original.Roots) {
		t.Fatal("copied graph closure")
	}
	for _, root := range original.Roots {
		found := false
		for _, next := range copied.Roots {
			if next.Role == root.Role && next.PartName == receipt.PartMap[root.PartName] {
				found = true
			}
		}
		if !found {
			t.Fatal("copied role root mapping", root)
		}
	}
	after := graphicsSessionMembers(t, target)
	for _, p := range original.Parts {
		mapped, ok := receipt.PartMap[p.PartName]
		if !ok || mapped == p.PartName {
			t.Fatal("copy part mapping")
		}
		if _, exists := targetBefore[mapped]; exists {
			t.Fatal("copy collided")
		}
		if strings.HasPrefix(p.ContentType, "image/") && !bytes.Equal(after[mapped], sourceBefore[p.PartName]) {
			t.Fatal("copied media custody")
		}
	}
	for _, edge := range original.Edges {
		owner := receipt.PartMap[edge.Owner]
		if edge.Owner == sourcePart {
			owner = targetPart
		}
		found := false
		for _, copyEdge := range copied.Edges {
			if copyEdge.Owner != owner || copyEdge.Type != edge.Type {
				continue
			}
			if edge.Owner != sourcePart && copyEdge.RelationshipID != edge.RelationshipID {
				continue
			}
			if edge.External && copyEdge.External && copyEdge.Target == edge.Target {
				found = true
			}
			if !edge.External && !copyEdge.External && copyEdge.PartName != nil && edge.PartName != nil && *copyEdge.PartName == receipt.PartMap[*edge.PartName] {
				found = true
			}
		}
		if !found {
			t.Fatal("copied edge retarget/custody", edge)
		}
	}
	for n, raw := range targetBefore {
		if n != targetPart && n != packaging.RelationshipsPathForPart(targetPart) && n != packaging.ContentTypesPath && !bytes.Equal(after[n], raw) {
			t.Fatal("target custody", n)
		}
	}
	currentSource := graphicsSessionMembers(t, source)
	for n, raw := range sourceBefore {
		if !same || n != sourcePart && n != packaging.RelationshipsPathForPart(sourcePart) && n != packaging.ContentTypesPath {
			if !bytes.Equal(currentSource[n], raw) {
				t.Fatal("source custody", n)
			}
		}
	}
	for old, next := range receipt.ModelIDMap {
		if old == next || !strings.HasPrefix(next, "{") {
			t.Fatal("fresh model identity")
		}
	}
	var out bytes.Buffer
	if e = target.pkg.WriteTo(&out); e != nil {
		t.Fatal(e)
	}
	reopened, e := OpenEditing(out.Bytes(), packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	read, e := reopened.InspectSmartArt(targetPart)
	if e != nil || !reflect.DeepEqual(read, b) {
		t.Fatal("copied reopen graph", e)
	}
}
func TestGraphicsSmartArtCopyRecipes(t *testing.T) {
	root, assets := graphicsRoot(t)
	var r struct {
		Feature           string `json:"feature"`
		Contract          string `json:"contract"`
		BaseFixtureID     string `json:"baseFixtureId"`
		SlidePart         string `json:"slidePart"`
		DestinationRecipe struct {
			Title string `json:"title"`
		} `json:"destinationRecipe"`
		Cases []struct {
			ScenarioID string              `json:"scenarioId"`
			CaseID     string              `json:"caseId"`
			Operations []graphicsOperation `json:"operations"`
			Request    struct {
				ShapeID            uint32 `json:"shapeId"`
				SameSlide          bool   `json:"sameSlide"`
				ProtectDestination bool   `json:"protectDestination"`
			} `json:"request"`
			Expected struct {
				ShapeID           uint32 `json:"shapeId"`
				Parts             int    `json:"parts"`
				NodeIdentities    int    `json:"nodeIdentityCount"`
				DrawingIdentities int    `json:"drawingIdentityCount"`
			} `json:"expected"`
			ErrorCode string `json:"errorCode"`
		} `json:"cases"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-smartart-copy.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, r.Feature)
	graphicsReadAsset(t, root, assets, r.Contract)
	base := graphicsReadAsset(t, root, assets, r.BaseFixtureID)
	if len(r.Cases) != 8 {
		t.Fatal("copy inventory")
	}
	for _, c := range r.Cases {
		t.Run(c.ScenarioID+" ["+c.CaseID+"]", func(t *testing.T) {
			input := graphicsInput(t, base, c.Operations)
			source, e := OpenEditing(input, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target := source
			if !c.Request.SameSlide {
				target = graphicsCopyTarget(t, r.DestinationRecipe.Title)
			}
			part := "ppt/slides/slide1.xml"
			if c.Request.ProtectDestination {
				raw, h, er := target.pkg.Part(packaging.PresentationPath)
				if er != nil {
					t.Fatal(er)
				}
				next := []byte(strings.Replace(string(raw), "</p:presentation>", "<p:modifyVerifier/></p:presentation>", 1))
				if er = target.pkg.Replace([]packaging.Replacement{{Part: packaging.PresentationPath, ExpectedSHA256: h, Data: next}}); er != nil {
					t.Fatal(er)
				}
			}
			sb, tb := graphicsSessionMembers(t, source), graphicsSessionMembers(t, target)
			version := target.generation
			receipt, e := target.CopySmartArtFrom(part, source, r.SlidePart, c.Request.ShapeID)
			if c.ErrorCode != "" {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != c.ErrorCode {
					t.Fatalf("copy refusal %v want %s", e, c.ErrorCode)
				}
				if target.generation != version || !reflect.DeepEqual(tb, graphicsSessionMembers(t, target)) || !reflect.DeepEqual(sb, graphicsSessionMembers(t, source)) {
					t.Fatal("copy refusal custody")
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			if receipt.ShapeID != c.Expected.ShapeID || receipt.PartName != part || len(receipt.PartMap) != c.Expected.Parts || len(receipt.ModelIDMap) != c.Expected.NodeIdentities || len(receipt.DrawingIDMap) != c.Expected.DrawingIdentities || target.generation != version+1 {
				t.Fatalf("copy receipt %#v", receipt)
			}
			graphicsVerifySmartCopy(t, source, target, r.SlidePart, part, c.Request.ShapeID, receipt, sb, tb, c.Request.SameSlide)
			if receipt.DrawingIDMap["ppt/diagrams/drawing1.xml#5"] != 6 {
				t.Fatal("synthetic drawing remap")
			}
		})
	}
}
func TestGraphicsSmartArtSealedSourceCopies(t *testing.T) {
	root, assets := graphicsRoot(t)
	var r struct {
		FixtureID string `json:"fixtureId"`
	}
	if e := json.Unmarshal(graphicsReadAsset(t, root, assets, "ledgers/pptx-smartart-office-source.json"), &r); e != nil {
		t.Fatal(e)
	}
	graphicsReadAsset(t, root, assets, "workflows/pptx/smartart-office-source.feature")
	graphicsReadAsset(t, root, assets, "contracts/pptx-smartart-office-source.md")
	input := graphicsReadAsset(t, root, assets, r.FixtureID)
	original := bytes.Clone(input)
	for _, same := range []bool{false, true} {
		t.Run("@id-pptx-smartart-office-source-copy ["+map[bool]string{false: "cross-presentation", true: "same-slide"}[same]+"]", func(t *testing.T) {
			source, e := OpenEditing(input, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target := source
			if !same {
				target = graphicsCopyTarget(t, "Copied source SmartArt")
			}
			part := "ppt/slides/slide1.xml"
			sb, tb := graphicsSessionMembers(t, source), graphicsSessionMembers(t, target)
			version := target.generation
			receipt, e := target.CopySmartArtFrom(part, source, part, 4)
			if e != nil {
				t.Fatal(e)
			}
			maxID := uint64(0)
			targetDoc, err := losslessxml.Parse(tb[part])
			if err != nil {
				t.Fatal(err)
			}
			for _, n := range targetDoc.Elements() {
				if n.Name() == name(packaging.NSPresentationML, "cNvPr") {
					raw, _ := graphicsAttr(n, "", "id")
					v, err := strconv.ParseUint(raw, 10, 32)
					if err != nil {
						t.Fatal(err)
					}
					if v > maxID {
						maxID = v
					}
				}
			}
			if uint64(receipt.ShapeID) != maxID+1 || target.generation != version+1 || !bytes.Equal(input, original) || len(receipt.PartMap) != 5 || len(receipt.ModelIDMap) != 46 || len(receipt.DrawingIDMap) != 6 {
				t.Fatal("source copy maps")
			}
			graphicsVerifySmartCopy(t, source, target, part, part, 4, receipt, sb, tb, same)
			for _, p := range []string{"ppt/diagrams/layout1.xml", "ppt/diagrams/quickStyle1.xml", "ppt/diagrams/colors1.xml"} {
				raw, _, e := target.pkg.Part(receipt.PartMap[p])
				if e != nil || !bytes.Equal(raw, sb[p]) {
					t.Fatal("template bytes changed", p, e)
				}
			}
			data, _, e := target.pkg.Part(receipt.PartMap["ppt/diagrams/data1.xml"])
			if e != nil {
				t.Fatal(e)
			}
			doc, e := losslessxml.Parse(data)
			if e != nil {
				t.Fatal(e)
			}
			rid := ""
			for _, n := range doc.Elements() {
				if n.Name() == name(graphicsPersistDiagramNS, "dataModelExt") {
					rid, _ = graphicsAttr(n, "", "relId")
				}
			}
			g, e := target.pkg.Graph()
			if e != nil {
				t.Fatal(e)
			}
			found := false
			for _, edge := range g.Edges {
				if edge.Source == part && edge.ID == rid && edge.Type == graphicsDiagramDrawingRel && edge.ResolvedPart == receipt.PartMap["ppt/diagrams/drawing1.xml"] {
					found = true
				}
			}
			if !found {
				t.Fatal("copied slide drawing binding")
			}
			graphicsAssertSealedSmartCopy(t, target, part, receipt, sb, tb, rid)
			var out bytes.Buffer
			if e = target.pkg.WriteTo(&out); e != nil {
				t.Fatal(e)
			}
			filename := "smartart-cross-presentation.pptx"
			if same {
				filename = "smartart-same-slide.pptx"
			}
			graphicsWriteOutput(t, root, filename, out.Bytes())
		})
	}
	graphicsWriteOutput(t, root, "smartart-source.pptx", input)
}

// Compare parsed attributes independently, then mask only their original value
// spans. Template IDs, URIs, text, prefixes, whitespace and quotes stay literal.
func graphicsAssertSealedSmartCopy(t *testing.T, target *EditSession, part string, receipt SmartArtCopyReceipt, sourceBefore, targetBefore map[string][]byte, drawingRID string) {
	t.Helper()
	after := graphicsSessionMembers(t, target)
	if len(after) != len(targetBefore)+5 {
		t.Fatal("exact sealed copy membership")
	}
	fresh := map[string]bool{}
	for old, next := range receipt.ModelIDMap {
		if !graphicsGUID.MatchString(next) || fresh[next] || old == next {
			t.Fatal("unique fresh model GUID", old, next)
		}
		fresh[next] = true
	}
	refs := map[string]bool{}
	drawings := map[uint32]bool{}
	for _, p := range []string{"ppt/diagrams/data1.xml", "ppt/diagrams/drawing1.xml"} {
		old, e := losslessxml.Parse(sourceBefore[p])
		if e != nil {
			t.Fatal(e)
		}
		next, e := losslessxml.Parse(after[receipt.PartMap[p]])
		if e != nil {
			t.Fatal(e)
		}
		a, b := old.Elements(), next.Elements()
		if len(a) != len(b) {
			t.Fatal("instance element topology")
		}
		oldEdits, newEdits := []losslessxml.AttributeEdit{}, []losslessxml.AttributeEdit{}
		drawingIndex := 0
		for i, n := range a {
			if n.Name() != b[i].Name() {
				t.Fatal("instance element identity")
			}
			attrs := n.Attributes()
			if len(attrs) != len(b[i].Attributes()) {
				t.Fatal("instance attribute topology")
			}
			for j, attr := range attrs {
				want := attr.Value
				allowed := false
				if attr.Name.Space == "" && graphicsChoice(attr.Name.Local, []string{"modelId", "srcId", "destId", "cxnId", "parTransId", "sibTransId", "presId", "presAssocID"}) {
					if mapped, ok := receipt.ModelIDMap[attr.Value]; ok {
						want = mapped
						allowed = true
						refs[attr.Value] = true
					}
				}
				if n.Name() == name(graphicsPersistDiagramNS, "cNvPr") && attr.Name == name("", "id") {
					key := p + "#0@" + strconv.Itoa(drawingIndex)
					drawingIndex++
					mapped, ok := receipt.DrawingIDMap[key]
					if !ok || mapped == 0 || drawings[mapped] {
						t.Fatal("zero drawing identity isolation", key)
					}
					drawings[mapped] = true
					want = strconv.FormatUint(uint64(mapped), 10)
					allowed = true
				}
				if n.Name() == name(graphicsPersistDiagramNS, "dataModelExt") && attr.Name == name("", "relId") {
					want = drawingRID
					allowed = true
				}
				got := b[i].Attributes()[j]
				if got.Name != attr.Name || got.Value != want {
					t.Fatal("copied instance attribute", p, n.Name(), attr.Name, got.Value, want)
				}
				if allowed {
					oldEdits = append(oldEdits, losslessxml.AttributeEdit{Target: n, Name: attr.Name, Value: "MASK"})
					newEdits = append(newEdits, losslessxml.AttributeEdit{Target: b[i], Name: attr.Name, Value: "MASK"})
				}
			}
		}
		maskedOld, e := old.Edit(nil, oldEdits)
		if e != nil {
			t.Fatal(e)
		}
		maskedNew, e := next.Edit(nil, newEdits)
		if e != nil {
			t.Fatal(e)
		}
		if !bytes.Equal(maskedOld, maskedNew) {
			t.Fatal("instance XML changed outside remapped scalar spans", p)
		}
	}
	if len(refs) != 46 || len(drawings) != 6 {
		t.Fatal("all sealed instance identities must be checked")
	}
	slide, e := losslessxml.Parse(after[part])
	if e != nil {
		t.Fatal(e)
	}
	var copied []losslessxml.Element
	for _, n := range slide.Elements() {
		if n.Name() != name(packaging.NSPresentationML, "graphicFrame") {
			continue
		}
		nv, e := graphicsOne(slide, n, packaging.NSPresentationML, "nvGraphicFramePr", false)
		if e != nil {
			t.Fatal(e)
		}
		props, e := graphicsOne(slide, nv, packaging.NSPresentationML, "cNvPr", false)
		if e != nil {
			t.Fatal(e)
		}
		id, _ := graphicsAttr(props, "", "id")
		if id == strconv.FormatUint(uint64(receipt.ShapeID), 10) {
			copied = append(copied, n)
		}
	}
	if len(copied) != 1 {
		t.Fatal("one appended frame")
	}
	unchanged, e := slide.RemoveElements(copied)
	if e != nil || !bytes.Equal(unchanged, targetBefore[part]) {
		t.Fatal("destination literal frame/sibling custody", e)
	}
	for _, registry := range []string{packaging.RelationshipsPathForPart(part), packaging.ContentTypesPath} {
		d, e := losslessxml.Parse(after[registry])
		if e != nil {
			t.Fatal(e)
		}
		remove := []losslessxml.Element{}
		original, e := losslessxml.Parse(targetBefore[registry])
		if e != nil {
			t.Fatal(e)
		}
		known := map[string]bool{}
		key := "Id"
		if registry == packaging.ContentTypesPath {
			key = "PartName"
		}
		for _, n := range original.Elements()[1:] {
			v, _ := graphicsAttr(n, "", key)
			known[v] = true
		}
		for _, n := range d.Elements()[1:] {
			v, _ := graphicsAttr(n, "", key)
			if !known[v] {
				remove = append(remove, n)
			}
		}
		if len(remove) != 5 {
			t.Fatal("exactly five registry additions", registry, len(remove))
		}
		retained, e := d.RemoveElements(remove)
		if e != nil || !bytes.Equal(retained, targetBefore[registry]) {
			t.Fatal("unrelated registry lexical custody", registry, e)
		}
	}
}
