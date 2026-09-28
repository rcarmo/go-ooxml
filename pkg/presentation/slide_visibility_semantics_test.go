package presentation

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// This verifies Go reader semantics. The retained S archive has a namespaced
// p:show marker, not an application-confirmed hidden slide.
func TestSlideHiddenAttributeQualification(t *testing.T) {
	fixture, err := testutil.LookupFixture("fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c")
	if err != nil {
		t.Fatal(err)
	}
	original, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		name, old, replacement string
		wantHidden             bool
		wantSavedShow          string
	}{
		{"retained-namespaced", "", "", false, ""},
		{"unqualified-hidden", `p:show="0"`, `show="0"`, true, "false"},
		{"unqualified-visible", `p:show="0"`, `show="1"`, false, "true"},
		{"absent-default", ` p:show="0"`, "", false, ""},
		{"both-unqualified-hidden", `p:show="0"`, `p:show="1" show="0"`, true, "false"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			data := original
			if tc.old != "" {
				data = replaceRetainedZipMember(t, original, "ppt/slides/slide3.xml", tc.old, tc.replacement)
			}
			readSlideHiddenStates(t, data, tc.wantHidden)

			// SaveAs serializes all slide parts. It must retain the schema
			// visibility value and never promote p:show to show="0".
			p, err := OpenReader(bytes.NewReader(data), int64(len(data)))
			if err != nil {
				t.Fatal(err)
			}
			outputPath := filepath.Join(t.TempDir(), "reopened.pptx")
			if err := p.SaveAs(outputPath); err != nil {
				t.Fatal(err)
			}
			output, err := os.ReadFile(outputPath)
			if err != nil {
				t.Fatal(err)
			}
			readSlideHiddenStates(t, output, tc.wantHidden)
			root := slide3RootAttributes(t, output)
			var unqualifiedShow string
			for _, a := range root {
				if a.Name.Space == "" && a.Name.Local == "show" {
					unqualifiedShow = a.Value
				}
			}
			if unqualifiedShow != tc.wantSavedShow {
				t.Fatalf("saved slide 3 unqualified show = %q, want %q", unqualifiedShow, tc.wantSavedShow)
			}
		})
	}
	after, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(original, after) {
		t.Fatal("retained source fixture changed")
	}
}

func readSlideHiddenStates(t *testing.T, data []byte, thirdHidden bool) {
	t.Helper()
	p, err := OpenReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	slides := p.Slides()
	if len(slides) != 4 {
		t.Fatalf("loaded %d slides, want four", len(slides))
	}
	for i, slide := range slides {
		want := i == 2 && thirdHidden
		if slide.Hidden() != want {
			t.Fatalf("slide %d Hidden() = %v, want %v", i+1, slide.Hidden(), want)
		}
	}
}

func slide3RootAttributes(t *testing.T, data []byte) []xml.Attr {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range zr.File {
		if file.Name != "ppt/slides/slide3.xml" {
			continue
		}
		r, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		dec := xml.NewDecoder(r)
		for {
			tok, err := dec.Token()
			if err != nil {
				r.Close()
				t.Fatal(err)
			}
			if root, ok := tok.(xml.StartElement); ok {
				r.Close()
				if root.Name != (xml.Name{Space: pmlNamespace, Local: "sld"}) {
					t.Fatalf("slide 3 root = %v", root.Name)
				}
				return root.Attr
			}
		}
	}
	t.Fatal("missing slide 3")
	return nil
}
