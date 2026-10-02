package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"strconv"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

const retainedVisibilityFixture = "fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c"
const visibleControlFixture = "fixture-fa245a3df00fef7f7bf4739921ee840194040161e06490589e3d52cc9fa7a71d"

func retainedVisibilityEvidenceSteps(sc *godog.ScenarioContext) {
	var original, control, sourceCopy, controlCopy []byte
	var sourceSlides, controlSlides []presentation.RetainedSlideEvidence
	var sourceRaw, controlRaw []rawVisibilitySlide
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		original, control, sourceCopy, controlCopy = nil, nil, nil, nil
		sourceSlides, controlSlides = nil, nil
		sourceRaw, controlRaw = nil, nil
		return ctx, nil
	})
	load := func(id string) ([]byte, error) {
		path, err := testutil.LookupFixture(id)
		if err != nil {
			return nil, err
		}
		return os.ReadFile(path)
	}
	sc.Step(`^the committed fixture (fixture-[0-9a-f]+) is loaded without editing$`, func(id string) error {
		if id != retainedVisibilityFixture {
			return fmt.Errorf("unexpected retained fixture %s", id)
		}
		var err error
		original, err = load(id)
		sourceCopy = bytes.Clone(original)
		return err
	})
	sc.Step(`^the committed fixture (fixture-[0-9a-f]+) is loaded as a visible control$`, func(id string) error {
		if id != visibleControlFixture {
			return fmt.Errorf("unexpected control fixture %s", id)
		}
		var err error
		control, err = load(id)
		controlCopy = bytes.Clone(control)
		return err
	})
	sc.Step(`^their four ordered slide identities and root visibility attributes are inspected$`, func() error {
		if original == nil || control == nil {
			return fmt.Errorf("missing retained inputs")
		}
		var err error
		sourceSlides, err = presentation.InspectRetainedSlideVisibility(original)
		if err != nil {
			return err
		}
		controlSlides, err = presentation.InspectRetainedSlideVisibility(control)
		if err != nil {
			return err
		}
		sourceRaw, err = inspectRawVisibility(original)
		if err != nil {
			return err
		}
		controlRaw, err = inspectRawVisibility(control)
		return err
	})
	sc.Step(`^the source fixture has exactly four slides with slide 3 marked namespaced p:show="0" and slides 1, 2 and 4 unmarked$`, func() error {
		if len(sourceSlides) != 4 || len(sourceRaw) != 4 {
			return fmt.Errorf("source slide count %d/%d", len(sourceSlides), len(sourceRaw))
		}
		for i, slide := range sourceSlides {
			want := ""
			if i == 2 {
				want = "0"
			}
			if slide.NamespacedShow != want || sourceRaw[i].marker != want {
				return fmt.Errorf("slide %d marker %q, raw %q, want %q", i+1, slide.NamespacedShow, sourceRaw[i].marker, want)
			}
		}
		return nil
	})
	sc.Step(`^the source fixture's four slides lack an unqualified show attribute$`, func() error {
		if len(sourceRaw) != 4 {
			return fmt.Errorf("source raw slide inspection absent")
		}
		for i, slide := range sourceRaw {
			if slide.unqualifiedShow {
				return fmt.Errorf("source slide %d has unqualified show", i+1)
			}
		}
		return nil
	})
	sc.Step(`^the visible control has exactly four slides without a hidden visibility attribute$`, func() error {
		if len(controlSlides) != 4 || len(controlRaw) != 4 {
			return fmt.Errorf("control slide count %d/%d", len(controlSlides), len(controlRaw))
		}
		for i, slide := range controlSlides {
			if slide.NamespacedShow != "" || controlRaw[i].marker != "" || controlRaw[i].unqualifiedShow {
				return fmt.Errorf("control slide %d marker %q, raw %+v", i+1, slide.NamespacedShow, controlRaw[i])
			}
		}
		return nil
	})
	sc.Step(`^no slide in either retained input is counted as an application-confirmed hidden slide$`, func() error {
		if len(sourceSlides) != 4 || len(controlSlides) != 4 {
			return fmt.Errorf("missing package inspection")
		}
		for _, group := range [][]presentation.RetainedSlideEvidence{sourceSlides, controlSlides} {
			for _, slide := range group {
				if slide.ApplicationConfirmedHidden {
					return fmt.Errorf("unsupported application-confirmed hidden status")
				}
			}
		}
		return nil
	})
	sc.Step(`^both source archive bytes remain unchanged$`, func() error {
		if !bytes.Equal(original, sourceCopy) || !bytes.Equal(control, controlCopy) {
			return fmt.Errorf("caller input mutated")
		}
		return nil
	})
	sc.Step(`^the reported slide numbers refer to the same slide relationships as the inspected parts$`, func() error {
		if err := compareRawVisibility(sourceSlides, sourceRaw, [4]string{"rId7", "rId8", "rId9", "rId10"}); err != nil {
			return err
		}
		return compareRawVisibility(controlSlides, controlRaw, [4]string{"rId2", "rId3", "rId4", "rId5"})
	})
}

// This acceptance oracle reads ZIP/XML independently of presentation's inspector.
// It follows the presentation's slide-ID list through relationship IDs to parts.
type rawVisibilitySlide struct {
	id                uint64
	rid, part, marker string
	unqualifiedShow   bool
}

const rawP = "http://schemas.openxmlformats.org/presentationml/2006/main"
const rawR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const rawPackageR = "http://schemas.openxmlformats.org/package/2006/relationships"

func rawXML(data []byte) ([]xml.StartElement, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var starts []xml.StartElement
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return starts, nil
		}
		if err != nil {
			return nil, err
		}
		if x, ok := tok.(xml.StartElement); ok {
			starts = append(starts, x)
		}
	}
}
func rawAttr(n xml.StartElement, namespace, local string) string {
	for _, a := range n.Attr {
		if a.Name == (xml.Name{Space: namespace, Local: local}) {
			return a.Value
		}
	}
	return ""
}
func inspectRawVisibility(source []byte) ([]rawVisibilitySlide, error) {
	zr, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		return nil, err
	}
	members := map[string][]byte{}
	for _, f := range zr.File {
		if _, ok := members[f.Name]; ok {
			return nil, fmt.Errorf("duplicate member %s", f.Name)
		}
		if f.UncompressedSize64 > 8<<20 {
			return nil, fmt.Errorf("oversize member %s", f.Name)
		}
		r, e := f.Open()
		if e != nil {
			return nil, e
		}
		b, e := io.ReadAll(io.LimitReader(r, (8<<20)+1))
		r.Close()
		if e != nil {
			return nil, e
		}
		if len(b) > 8<<20 {
			return nil, fmt.Errorf("oversize member %s", f.Name)
		}
		members[f.Name] = b
	}
	ids, err := rawXML(members["ppt/presentation.xml"])
	if err != nil {
		return nil, err
	}
	slides := []rawVisibilitySlide{}
	for _, n := range ids {
		if n.Name == (xml.Name{Space: rawP, Local: "sldId"}) {
			id, e := strconv.ParseUint(rawAttr(n, "", "id"), 10, 64)
			if e != nil {
				return nil, e
			}
			slides = append(slides, rawVisibilitySlide{id: id, rid: rawAttr(n, rawR, "id")})
		}
	}
	rels, err := rawXML(members["ppt/_rels/presentation.xml.rels"])
	if err != nil {
		return nil, err
	}
	targets := map[string]string{}
	for _, n := range rels {
		if n.Name == (xml.Name{Space: rawPackageR, Local: "Relationship"}) && rawAttr(n, "", "Type") == rawR+"/slide" {
			if rawAttr(n, "", "TargetMode") != "" {
				return nil, fmt.Errorf("external slide relationship")
			}
			id := rawAttr(n, "", "Id")
			if _, ok := targets[id]; ok {
				return nil, fmt.Errorf("duplicate slide relationship")
			}
			targets[id] = rawAttr(n, "", "Target")
		}
	}
	if len(slides) != 4 || len(targets) != 4 {
		return nil, fmt.Errorf("raw slide/relationship count %d/%d", len(slides), len(targets))
	}
	for i := range slides {
		target := targets[slides[i].rid]
		if target == "" || strings.HasPrefix(target, "/") || strings.Contains(target, "\\") || strings.Contains(target, ":") {
			return nil, fmt.Errorf("missing/unsafe slide relationship %d", i+1)
		}
		slides[i].part = path.Clean(path.Join("ppt", target))
		if !strings.HasPrefix(slides[i].part, "ppt/slides/") {
			return nil, fmt.Errorf("escaping slide part")
		}
		nodes, e := rawXML(members[slides[i].part])
		if e != nil {
			return nil, e
		}
		if len(nodes) == 0 || nodes[0].Name != (xml.Name{Space: rawP, Local: "sld"}) {
			return nil, fmt.Errorf("slide %d root", i+1)
		}
		for _, a := range nodes[0].Attr {
			if a.Name.Local == "show" {
				if a.Name.Space == "" {
					slides[i].unqualifiedShow = true
				} else if a.Name.Space == rawP {
					slides[i].marker = a.Value
				} else {
					return nil, fmt.Errorf("foreign slide marker")
				}
			}
		}
	}
	return slides, nil
}
func compareRawVisibility(got []presentation.RetainedSlideEvidence, raw []rawVisibilitySlide, rids [4]string) error {
	if len(got) != 4 || len(raw) != 4 {
		return fmt.Errorf("incomplete evidence %d/%d", len(got), len(raw))
	}
	for i := 0; i < 4; i++ {
		wantPart := fmt.Sprintf("ppt/slides/slide%d.xml", i+1)
		if got[i].SlideID != uint64(256+i) || raw[i].id != uint64(256+i) || got[i].RelationshipID != rids[i] || raw[i].rid != rids[i] || got[i].Part != wantPart || raw[i].part != wantPart || got[i].NamespacedShow != raw[i].marker || raw[i].unqualifiedShow {
			return fmt.Errorf("slide %d mismatch inspector=%+v raw=%+v", i+1, got[i], raw[i])
		}
	}
	return nil
}

func TestRetainedVisibilityEvidenceOracleRejectsCorruption(t *testing.T) {
	path, err := testutil.LookupFixture(retainedVisibilityFixture)
	if err != nil {
		t.Fatal(err)
	}
	source, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	base, err := inspectRawVisibility(source)
	if err != nil {
		t.Fatal(err)
	}
	got, err := presentation.InspectRetainedSlideVisibility(source)
	if err != nil {
		t.Fatal(err)
	}
	cases := []struct {
		name  string
		alter func([]rawVisibilitySlide) []rawVisibilitySlide
	}{
		{"unqualified marker", func(x []rawVisibilitySlide) []rawVisibilitySlide { x[2].unqualifiedShow = true; return x }},
		{"wrong namespaced marker", func(x []rawVisibilitySlide) []rawVisibilitySlide { x[2].marker = "1"; return x }},
		{"wrong relationship", func(x []rawVisibilitySlide) []rawVisibilitySlide { x[2].rid = "rId10"; return x }},
		{"wrong part", func(x []rawVisibilitySlide) []rawVisibilitySlide { x[2].part = "ppt/slides/slide4.xml"; return x }},
		{"wrong ordinal", func(x []rawVisibilitySlide) []rawVisibilitySlide { x[2].id = 260; return x }},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			changed := tc.alter(append([]rawVisibilitySlide(nil), base...))
			if err := compareRawVisibility(got, changed, [4]string{"rId7", "rId8", "rId9", "rId10"}); err == nil {
				t.Fatal("independent corruption escaped")
			}
		})
	}
}
