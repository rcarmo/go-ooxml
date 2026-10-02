package presentation

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// TestRetainedSlideVisibilityStructure tests package evidence only. Neither
// fixture has PowerPoint-confirmed hidden-slide status.
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
