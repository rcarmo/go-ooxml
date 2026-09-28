package presentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"strconv"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

const (
	pmlNamespace                  = "http://schemas.openxmlformats.org/presentationml/2006/main"
	officeRelationshipsNamespace  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	packageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
	slideRelationshipType         = officeRelationshipsNamespace + "/slide"
)

type retainedSlide struct {
	id     uint64
	rid    string
	part   string
	marker string // Only the namespaced p:show marker is recorded; an unqualified show is refused.
}

type retainedRelationship struct {
	typ, target, mode string
}

// TestRetainedSlideVisibilityStructure tests package evidence only. Neither
// fixture has PowerPoint-confirmed hidden-slide status; this test does not run
// or select the shared slide-visibility Gherkin scenario.
func TestRetainedSlideVisibilityStructure(t *testing.T) {
	cases := []struct {
		name, asset string
		rids        [4]string
		marker      string
	}{
		{"retained-S", "fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c", [4]string{"rId7", "rId8", "rId9", "rId10"}, "0"},
		{"visible-P", "fixture-fa245a3df00fef7f7bf4739921ee840194040161e06490589e3d52cc9fa7a71d", [4]string{"rId2", "rId3", "rId4", "rId5"}, ""},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			fixture, err := testutil.LookupFixture(tc.asset)
			if err != nil {
				t.Fatal(err)
			}
			original, err := os.ReadFile(fixture)
			if err != nil {
				t.Fatal(err)
			}
			if err := verifyRetainedVisibility(original, expectedRetainedSlides(tc.rids, tc.marker)); err != nil {
				t.Fatal(err)
			}
			after, err := os.ReadFile(fixture)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(original, after) {
				t.Fatal("retained source archive changed during inspection")
			}
		})
	}
}

func TestRetainedSlideVisibilityRejectsBrokenStructure(t *testing.T) {
	fixture, err := testutil.LookupFixture("fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c")
	if err != nil {
		t.Fatal(err)
	}
	original, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	cases := []struct {
		name, part, from, to string
	}{
		{"unqualified-show", "ppt/slides/slide3.xml", `p:show="0"`, `show="0"`},
		{"wrong-show-namespace", "ppt/slides/slide3.xml", `p:show="0"`, `r:show="0"`},
		{"missing-marker", "ppt/slides/slide3.xml", ` p:show="0"`, ``},
		{"wrong-marker-slide", "ppt/slides/slide2.xml", `<p:sld `, `<p:sld p:show="0" `},
		{"duplicate-slide-relationship", "ppt/_rels/presentation.xml.rels", `Id="rId9"`, `Id="rId8"`},
		{"reused-slide-part", "ppt/_rels/presentation.xml.rels", `Target="slides/slide3.xml"`, `Target="slides/slide4.xml"`},
		{"escaping-slide-part", "ppt/_rels/presentation.xml.rels", `Target="slides/slide3.xml"`, `Target="../slide3.xml"`},
		{"external-slide-relationship", "ppt/_rels/presentation.xml.rels", `Target="slides/slide3.xml"`, `Target="https://example.invalid/slide3.xml" TargetMode="External"`},
		{"missing-slide-relationship", "ppt/presentation.xml", `r:id="rId9"`, `r:id="rId99"`},
		{"wrong-slide-relationship", "ppt/presentation.xml", `id="258" r:id="rId9"`, `id="258" r:id="rId10"`},
		{"duplicate-slide-id-list", "ppt/presentation.xml", `<p:sldIdLst>`, `<p:sldIdLst></p:sldIdLst><p:sldIdLst>`},
		{"reordered-slide-id", "ppt/presentation.xml", `id="258" r:id="rId9"`, `id="260" r:id="rId9"`},
		{"malformed-slide-xml", "ppt/slides/slide3.xml", `p:show="0"`, `p:show="0" p:show="1"`},
		{"duplicate-root-marker", "ppt/slides/slide3.xml", `p:show="0"`, `p:show="0" show="1"`},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			corrupt := replaceRetainedZipMember(t, original, tc.part, tc.from, tc.to)
			if err := verifyRetainedVisibility(corrupt, expectedRetainedSlides([4]string{"rId7", "rId8", "rId9", "rId10"}, "0")); err == nil {
				t.Fatal("accepted a changed visibility marker or slide graph")
			}
		})
	}
	// A valid slide relationship with a missing target must fail, rather than
	// silently shrinking the slide list as the public presentation reader can.
	t.Run("missing-slide-part", func(t *testing.T) {
		corrupt := replaceRetainedZipMember(t, original, "ppt/_rels/presentation.xml.rels", `Target="slides/slide3.xml"`, `Target="slides/missing.xml"`)
		if _, err := inspectRetainedVisibility(corrupt); err == nil {
			t.Fatal("accepted a missing slide part")
		}
	})

	control, err := testutil.LookupFixture("fixture-fa245a3df00fef7f7bf4739921ee840194040161e06490589e3d52cc9fa7a71d")
	if err != nil {
		t.Fatal(err)
	}
	controlBytes, err := os.ReadFile(control)
	if err != nil {
		t.Fatal(err)
	}
	t.Run("control-unqualified-show", func(t *testing.T) {
		corrupt := replaceRetainedZipMember(t, controlBytes, "ppt/slides/slide3.xml", `<p:sld `, `<p:sld show="0" `)
		if err := verifyRetainedVisibility(corrupt, expectedRetainedSlides([4]string{"rId2", "rId3", "rId4", "rId5"}, "")); err == nil {
			t.Fatal("accepted an unqualified show attribute in the control")
		}
	})
}

func verifyRetainedVisibility(data []byte, want []retainedSlide) error {
	got, err := inspectRetainedVisibility(data)
	if err != nil {
		return err
	}
	if !equalRetainedSlides(got, want) {
		return fmt.Errorf("ordered slide identities/attributes: got %+v, want %+v", got, want)
	}
	return nil
}

func expectedRetainedSlides(rids [4]string, marker string) []retainedSlide {
	out := make([]retainedSlide, 4)
	for i := range out {
		out[i] = retainedSlide{id: uint64(256 + i), rid: rids[i], part: fmt.Sprintf("ppt/slides/slide%d.xml", i+1)}
	}
	out[2].marker = marker
	return out
}

func equalRetainedSlides(a, b []retainedSlide) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func inspectRetainedVisibility(data []byte) ([]retainedSlide, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	members := make(map[string]*zip.File, len(zr.File))
	for _, member := range zr.File {
		if _, duplicate := members[member.Name]; duplicate {
			return nil, fmt.Errorf("duplicate ZIP member %q", member.Name)
		}
		members[member.Name] = member
	}
	read := func(name string) ([]byte, error) {
		member, ok := members[name]
		if !ok || member.FileInfo().IsDir() || member.UncompressedSize64 > 8<<20 {
			return nil, fmt.Errorf("missing or oversized XML part %q", name)
		}
		r, err := member.Open()
		if err != nil {
			return nil, err
		}
		defer r.Close()
		content, err := io.ReadAll(io.LimitReader(r, 8<<20+1))
		if err != nil || len(content) > 8<<20 {
			return nil, fmt.Errorf("invalid or oversized XML part %q: %v", name, err)
		}
		return content, nil
	}
	presentationXML, err := read("ppt/presentation.xml")
	if err != nil {
		return nil, err
	}
	slides, err := retainedSlideIDs(presentationXML)
	if err != nil {
		return nil, err
	}
	relsXML, err := read("ppt/_rels/presentation.xml.rels")
	if err != nil {
		return nil, err
	}
	rels, slideRels, err := retainedRelationships(relsXML)
	if err != nil {
		return nil, err
	}
	if len(slides) != 4 || slideRels != 4 {
		return nil, fmt.Errorf("expected four slide IDs and relationships, got %d and %d", len(slides), slideRels)
	}
	seenParts := make(map[string]bool, 4)
	for i := range slides {
		rel, ok := rels[slides[i].rid]
		if !ok || rel.typ != slideRelationshipType || rel.mode != "" || rel.target == "" || strings.HasPrefix(rel.target, "/") || strings.Contains(rel.target, "\\") || strings.Contains(rel.target, ":") {
			return nil, fmt.Errorf("slide %d has no unique internal slide relationship", i+1)
		}
		part := path.Clean(path.Join("ppt", rel.target))
		if !strings.HasPrefix(part, "ppt/slides/") || seenParts[part] {
			return nil, fmt.Errorf("slide %d has escaping or reused part %q", i+1, part)
		}
		seenParts[part] = true
		slides[i].part = part
		content, err := read(part)
		if err != nil {
			return nil, err
		}
		slides[i].marker, err = retainedSlideRoot(content)
		if err != nil {
			return nil, fmt.Errorf("slide %d (%s): %w", i+1, part, err)
		}
	}
	return slides, nil
}

func retainedSlideIDs(data []byte) ([]retainedSlide, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	if _, err := retainedRoot(dec, xml.Name{Space: pmlNamespace, Local: "presentation"}); err != nil {
		return nil, err
	}
	var slides []retainedSlide
	depth, lists, listDepth := 1, 0, 0
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return nil, err
		}
		switch x := tok.(type) {
		case xml.StartElement:
			if depth == 1 && x.Name == (xml.Name{Space: pmlNamespace, Local: "sldIdLst"}) {
				lists++
				listDepth = depth + 1
			} else if listDepth != 0 && depth == listDepth && lists == 1 && x.Name == (xml.Name{Space: pmlNamespace, Local: "sldId"}) {
				var id, rid string
				for _, a := range x.Attr {
					switch a.Name {
					case xml.Name{Local: "id"}:
						if id != "" {
							return nil, fmt.Errorf("duplicate slide ID attribute")
						}
						id = a.Value
					case xml.Name{Space: officeRelationshipsNamespace, Local: "id"}:
						if rid != "" {
							return nil, fmt.Errorf("duplicate slide relationship attribute")
						}
						rid = a.Value
					}
				}
				n, err := strconv.ParseUint(id, 10, 64)
				if err != nil || rid == "" {
					return nil, fmt.Errorf("invalid slide ID or relationship")
				}
				slides = append(slides, retainedSlide{id: n, rid: rid})
			}
			depth++
		case xml.EndElement:
			if depth == listDepth && x.Name == (xml.Name{Space: pmlNamespace, Local: "sldIdLst"}) {
				listDepth = 0
			}
			depth--
		}
	}
	if err := retainedEOF(dec); err != nil || lists != 1 {
		return nil, fmt.Errorf("slide ID list or XML document invalid: %v", err)
	}
	seenIDs, seenRIDs := map[uint64]bool{}, map[string]bool{}
	for _, slide := range slides {
		if seenIDs[slide.id] || seenRIDs[slide.rid] {
			return nil, fmt.Errorf("duplicate slide identity")
		}
		seenIDs[slide.id], seenRIDs[slide.rid] = true, true
	}
	return slides, nil
}

func retainedRelationships(data []byte) (map[string]retainedRelationship, int, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	if _, err := retainedRoot(dec, xml.Name{Space: packageRelationshipsNamespace, Local: "Relationships"}); err != nil {
		return nil, 0, err
	}
	rels := map[string]retainedRelationship{}
	depth, slideCount := 1, 0
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return nil, 0, err
		}
		switch x := tok.(type) {
		case xml.StartElement:
			if depth == 1 && x.Name == (xml.Name{Space: packageRelationshipsNamespace, Local: "Relationship"}) {
				attrs := map[string]string{}
				for _, a := range x.Attr {
					if a.Name.Space != "" {
						return nil, 0, fmt.Errorf("qualified relationship attribute")
					}
					if _, duplicate := attrs[a.Name.Local]; duplicate {
						return nil, 0, fmt.Errorf("duplicate relationship attribute")
					}
					attrs[a.Name.Local] = a.Value
				}
				id := attrs["Id"]
				if _, duplicate := rels[id]; id == "" || duplicate {
					return nil, 0, fmt.Errorf("missing or duplicate relationship ID")
				}
				rels[id] = retainedRelationship{typ: attrs["Type"], target: attrs["Target"], mode: attrs["TargetMode"]}
				if attrs["Type"] == slideRelationshipType {
					slideCount++
				}
			}
			depth++
		case xml.EndElement:
			depth--
		}
	}
	if err := retainedEOF(dec); err != nil {
		return nil, 0, err
	}
	return rels, slideCount, nil
}

func retainedSlideRoot(data []byte) (string, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	root, err := retainedRoot(dec, xml.Name{Space: pmlNamespace, Local: "sld"})
	if err != nil {
		return "", err
	}
	marker := ""
	for _, a := range root.Attr {
		if a.Name.Local != "show" {
			continue
		}
		if a.Name.Space != pmlNamespace || marker != "" {
			return "", fmt.Errorf("unqualified, foreign or duplicate show attribute")
		}
		marker = a.Value
	}
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return "", err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return marker, retainedEOF(dec)
}

func retainedRoot(dec *xml.Decoder, want xml.Name) (xml.StartElement, error) {
	for {
		tok, err := dec.Token()
		if err != nil {
			return xml.StartElement{}, err
		}
		if root, ok := tok.(xml.StartElement); ok {
			if root.Name != want {
				return root, fmt.Errorf("root %v, want %v", root.Name, want)
			}
			return root, nil
		}
		if text, ok := tok.(xml.CharData); ok && len(bytes.TrimSpace(text)) != 0 {
			return xml.StartElement{}, fmt.Errorf("text before XML root")
		}
	}
}

func retainedEOF(dec *xml.Decoder) error {
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		if text, ok := tok.(xml.CharData); ok && len(bytes.TrimSpace(text)) == 0 {
			continue
		}
		if _, ok := tok.(xml.Comment); ok {
			continue
		}
		return fmt.Errorf("content after XML root")
	}
}

func replaceRetainedZipMember(t *testing.T, original []byte, name, from, to string) []byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(original), int64(len(original)))
	if err != nil {
		t.Fatal(err)
	}
	var b bytes.Buffer
	zw := zip.NewWriter(&b)
	found := false
	for _, file := range zr.File {
		r, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		content, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			t.Fatal(err)
		}
		if file.Name == name {
			if found || bytes.Count(content, []byte(from)) != 1 {
				t.Fatalf("mutation target %q missing or ambiguous", name)
			}
			content = bytes.Replace(content, []byte(from), []byte(to), 1)
			found = true
		}
		w, err := zw.Create(file.Name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(content); err != nil {
			t.Fatal(err)
		}
	}
	if !found {
		t.Fatalf("mutation member %q missing", name)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return b.Bytes()
}
