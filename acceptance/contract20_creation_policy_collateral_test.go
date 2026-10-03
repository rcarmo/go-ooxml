package acceptance

import (
	"bytes"
	"encoding/xml"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Native collateral only: these controls do not bind or credit any selected
// Contract20 creation scenario. A two-slide producer graph also exercises the
// policy's per-slide rather than aggregate edge requirement.
func TestContract20CreationPolicyCollateralBatch(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	c, err := presentation.NewContractPresentation()
	if err != nil {
		t.Fatal(err)
	}
	defer c.Close()
	for _, pair := range [][2]string{{"First", "One"}, {"Second", "Two"}} {
		if err := c.AddTitleSlide(pair[0], pair[1]); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "created.pptx")
	if err := c.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	archive, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	members, err := blankMemberPayloads(archive)
	if err != nil {
		t.Fatal(err)
	}
	graph, err := contract20TableGraph(archive)
	if err != nil {
		t.Fatal(err)
	}
	if err := contract20CreationPolicy(members, graph, 2); err != nil {
		t.Fatal(err)
	}
	reopened, err := presentation.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	if reopened.SlideCount() != 2 {
		t.Fatalf("two-slide reopen count %d", reopened.SlideCount())
	}
	for i, pair := range [][2]string{{"First", "One"}, {"Second", "Two"}} {
		slide, e := reopened.Slide(i + 1)
		if e != nil {
			t.Fatal(e)
		}
		values := []string{}
		for _, shape := range slide.Shapes() {
			values = append(values, shape.Text())
		}
		if !reflect.DeepEqual(values, []string{pair[0], pair[1]}) {
			t.Fatalf("slide %d title/subtitle = %q", i, values)
		}
	}
	clone := func() packaging.Graph {
		out := graph
		out.Parts = append([]packaging.GraphPart{}, graph.Parts...)
		out.Edges = append([]packaging.Edge{}, graph.Edges...)
		return out
	}
	t.Run("external-edge", func(t *testing.T) {
		g := clone()
		g.Edges[0].External = true
		if contract20CreationPolicy(members, g, 2) == nil {
			t.Fatal("external edge accepted")
		}
	})
	t.Run("unknown-edge", func(t *testing.T) {
		g := clone()
		g.Edges[0].Type = "urn:unknown"
		if contract20CreationPolicy(members, g, 2) == nil {
			t.Fatal("unknown edge accepted")
		}
	})
	t.Run("extra-role", func(t *testing.T) {
		g := clone()
		g.Parts[0].ContentType = "application/unknown"
		if contract20CreationPolicy(members, g, 2) == nil {
			t.Fatal("unknown role accepted")
		}
	})
	t.Run("lost-layout-edge", func(t *testing.T) {
		g := clone()
		for i, e := range g.Edges {
			if e.Type == packaging.RelTypeSlideLayout && e.Source != "" {
				g.Edges = append(g.Edges[:i], g.Edges[i+1:]...)
				break
			}
		}
		if contract20CreationPolicy(members, g, 2) == nil {
			t.Fatal("missing per-slide layout edge accepted")
		}
	})
	t.Run("orphan-registry", func(t *testing.T) {
		g := clone()
		g.Parts = append(g.Parts, packaging.GraphPart{Name: "ppt/_rels/orphan.xml.rels", ContentType: packaging.ContentTypeRelationships})
		m := make(map[string][]byte, len(members)+1)
		for k, v := range members {
			m[k] = v
		}
		m["ppt/_rels/orphan.xml.rels"] = []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
		if contract20CreationPolicy(m, g, 2) == nil {
			t.Fatal("orphan registry accepted")
		}
	})
	t.Run("wrong-layout-type", func(t *testing.T) {
		m := make(map[string][]byte, len(members))
		for k, v := range members {
			m[k] = v
		}
		for _, p := range graph.Parts {
			if p.ContentType != packaging.ContentTypeSlideLayout {
				continue
			}
			d, e := losslessxml.Parse(m[p.Name])
			if e != nil {
				t.Fatal(e)
			}
			root := d.Elements()[0]
			found := false
			for _, a := range root.Attributes() {
				if a.Name.Local == "type" && a.Name.Space == "" {
					found = true
				}
			}
			if !found {
				t.Fatal("layout type missing")
			}
			m[p.Name] = bytes.Replace(m[p.Name], []byte(`type="title"`), []byte(`type="blank"`), 1)
			if bytes.Equal(m[p.Name], members[p.Name]) {
				t.Fatal("layout mutation not applied")
			}
			break
		}
		if contract20CreationPolicy(m, graph, 2) == nil {
			t.Fatal("wrong layout type accepted")
		}
	})
	t.Run("visible-subtitle-and-extra-shape", func(t *testing.T) {
		geometry, err := presentation.NewContractPresentation()
		if err != nil {
			t.Fatal(err)
		}
		defer geometry.Close()
		if err = geometry.AddTitleSlide("Tables", ""); err != nil {
			t.Fatal(err)
		}
		table, err := geometry.AddTable(2, 3, 120, 240, 1001, 1003)
		if err != nil {
			t.Fatal(err)
		}
		if err = geometry.SetTableCellText(table, 0, 0, "  <Alpha & Beta>  "); err != nil {
			t.Fatal(err)
		}
		geometryPath := filepath.Join(t.TempDir(), "geometry.pptx")
		if err = geometry.SaveAs(geometryPath); err != nil {
			t.Fatal(err)
		}
		geometryArchive, err := os.ReadFile(geometryPath)
		if err != nil {
			t.Fatal(err)
		}
		geometryMembers, err := blankMemberPayloads(geometryArchive)
		if err != nil {
			t.Fatal(err)
		}
		geometryGraph, err := contract20TableGraph(geometryArchive)
		if err != nil {
			t.Fatal(err)
		}
		var slidePart string
		for _, p := range geometryGraph.Parts {
			if p.ContentType == packaging.ContentTypeSlide {
				slidePart = p.Name
				break
			}
		}
		if slidePart == "" {
			t.Fatal("created slide missing")
		}
		base := geometryMembers[slidePart]
		d, err := losslessxml.Parse(base)
		if err != nil {
			t.Fatal(err)
		}
		var subtitle losslessxml.Element
		found := 0
		for _, n := range d.Elements() {
			if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "ph"}) {
				continue
			}
			for _, a := range n.Attributes() {
				if a.Name.Space == "" && a.Name.Local == "type" && a.Value == "subTitle" {
					for ancestor, ok := n.Parent(); ok; ancestor, ok = ancestor.Parent() {
						if ancestor.Name() == (xml.Name{Space: packaging.NSPresentationML, Local: "sp"}) {
							subtitle = ancestor
							found++
							break
						}
					}
				}
			}
		}
		if found != 1 {
			t.Fatalf("subtitle owners %d", found)
		}
		start, end := subtitle.SourceRange()
		if start < 0 || end > len(base) || start >= end {
			t.Fatal("subtitle source span invalid")
		}
		var leaf losslessxml.Element
		leafCount := 0
		for _, n := range d.Elements() {
			if n.Name() != (xml.Name{Space: packaging.NSDrawingML, Local: "t"}) {
				continue
			}
			a, b := n.SourceRange()
			if a >= start && b <= end {
				leaf = n
				leafCount++
			}
		}
		if leafCount != 1 {
			t.Fatalf("subtitle text leaf count %d", leafCount)
		}
		oldText, ordinary := leaf.Text()
		if !ordinary || oldText != "" {
			t.Fatalf("geometry subtitle is not empty: %q", oldText)
		}
		changed, err := d.ReplaceText([]losslessxml.TextEdit{{Target: leaf, Text: "Extra subtitle"}})
		if err != nil {
			t.Fatalf("valid subtitle mutation refused: %v", err)
		}
		if _, err := losslessxml.Parse(changed); err != nil {
			t.Fatalf("subtitle negative must be valid XML: %v", err)
		}
		if err := contract20GeometryVisibleSlide(changed, "Tables", "  <Alpha & Beta>  "); err == nil {
			t.Fatal("extra subtitle text accepted")
		}
		// Duplicate a valid direct title shape as a new positive unique object;
		// its text is still unrequested and must be rejected by the oracle.
		var title losslessxml.Element
		found = 0
		for _, n := range d.Elements() {
			if n.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "ph"}) {
				continue
			}
			for _, a := range n.Attributes() {
				if a.Name.Space == "" && a.Name.Local == "type" && a.Value == "ctrTitle" {
					for ancestor, ok := n.Parent(); ok; ancestor, ok = ancestor.Parent() {
						if ancestor.Name() == (xml.Name{Space: packaging.NSPresentationML, Local: "sp"}) {
							title = ancestor
							found++
							break
						}
					}
				}
			}
		}
		if found != 1 {
			t.Fatalf("title owners %d", found)
		}
		first, last := title.SourceRange()
		raw := bytes.Clone(base[first:last])
		if !bytes.Contains(raw, []byte(`id="2"`)) {
			t.Fatal("title object ID not sealed for negative")
		}
		raw = bytes.Replace(raw, []byte(`id="2"`), []byte(`id="999"`), 1)
		var tree losslessxml.Element
		for _, n := range d.Elements() {
			if n.Name() == (xml.Name{Space: packaging.NSPresentationML, Local: "spTree"}) {
				tree = n
				break
			}
		}
		_, offset := tree.ContentRange()
		withShape := append(append(bytes.Clone(base[:offset]), raw...), base[offset:]...)
		if _, err := losslessxml.Parse(withShape); err != nil {
			t.Fatalf("extra shape negative must be valid XML: %v", err)
		}
		if err := contract20GeometryVisibleSlide(withShape, "Tables", "  <Alpha & Beta>  "); err == nil {
			t.Fatal("extra visible text shape accepted")
		}
	})
	t.Run("bad-title-atomic", func(t *testing.T) {
		before := c.SlideCount()
		e := c.AddTitleSlide("bad\x00", "ignored")
		var refusal *packaging.Refusal
		if !errors.As(e, &refusal) || refusal.Kind != "PPTX_ARGUMENT_INVALID" || c.SlideCount() != before {
			t.Fatalf("title refusal %v count %d", e, c.SlideCount())
		}
		next := filepath.Join(t.TempDir(), "after-refusal.pptx")
		if e = c.SaveAs(next); e != nil {
			t.Fatal(e)
		}
		data, e := os.ReadFile(next)
		if e != nil {
			t.Fatal(e)
		}
		payload, e := blankMemberPayloads(data)
		if e != nil {
			t.Fatal(e)
		}
		if !reflect.DeepEqual(payload, members) {
			t.Fatal("refusal changed creator output members")
		}
	})
}
