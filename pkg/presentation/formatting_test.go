package presentation

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const formattingFixtureID = "fixture-84e7a8a3681d0e43d4c8a29c37e68386721dfb5df13a36b8bd17a3ebc1fb1fb2"

func formattingFixture(t *testing.T) ([]byte, *EditSession) {
	t.Helper()
	path := testutil.ReferencePath("fixtures", "pptx", "formatting", "formatting-84e7a8a3681d.pptx")
	b, e := os.ReadFile(path)
	if os.IsNotExist(e) {
		t.Skip("sealed formatting candidate fixture absent")
	}
	if e != nil {
		t.Fatal(e)
	}
	h := sha256.Sum256(b)
	if hex.EncodeToString(h[:]) != strings.TrimPrefix(formattingFixtureID, "fixture-") {
		t.Fatal("candidate fixture provenance")
	}
	s, e := OpenEditing(b, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	return b, s
}
func formattingSave(t *testing.T, s *EditSession) []byte {
	t.Helper()
	p := filepath.Join(t.TempDir(), "format.pptx")
	_, e := s.SaveAs(p)
	if e != nil {
		t.Fatal(e)
	}
	b, e := os.ReadFile(p)
	if e != nil {
		t.Fatal(e)
	}
	return b
}
func formattingRefusal(t *testing.T, e error, kind string) {
	t.Helper()
	var r *packaging.Refusal
	if !errors.As(e, &r) || r.Kind != kind {
		t.Fatalf("refusal=%v want %s", e, kind)
	}
}
func formattingValue[T any](v T) *T { return &v }
func TestFormattingDirectProperties(t *testing.T) {
	cases := []struct {
		name       string
		run        int
		rp         RunFormatPatch
		pp         ParagraphFormatPatch
		geometry   GeometryPatch
		visibility *bool
		token      []byte
	}{
		{name: "hide", visibility: formattingValue(false), token: []byte(`show="0"`)},
		{name: "unhide", visibility: formattingValue(true), token: []byte(`show="1"`)},
		{name: "move", geometry: GeometryPatch{X: formattingValue(int64(914400)), Y: formattingValue(int64(1828800))}, token: []byte(`x="914400" y="1828800"`)},
		{name: "resize", geometry: GeometryPatch{Width: formattingValue(int64(3657600)), Height: formattingValue(int64(1828800))}, token: []byte(`cx="3657600" cy="1828800"`)},
		{name: "rotate", geometry: GeometryPatch{Rotation: formattingValue(int64(5400000)), FlipH: formattingValue(true), FlipV: formattingValue(false)}, token: []byte(`rot="5400000"`)},
		{name: "unbold-second", run: 1, rp: RunFormatPatch{Bold: formattingValue(false)}, token: []byte(`<a:rPr b="0"`)},
		{name: "italic", rp: RunFormatPatch{Italic: formattingValue(true)}, token: []byte(`i="1"`)},
		{name: "font", rp: RunFormatPatch{Latin: formattingValue("Aptos Display")}, token: []byte(`typeface="Aptos Display"`)},
		{name: "color", rp: RunFormatPatch{Color: formattingValue("A1B2C3")}, token: []byte(`val="A1B2C3"`)},
		{name: "remove-direct", rp: RunFormatPatch{RemoveDirect: true}, token: []byte(`lang="en-US"`)},
		{name: "align", pp: ParagraphFormatPatch{Alignment: formattingValue("ctr")}, token: []byte(`algn="ctr"`)},
		{name: "indent", pp: ParagraphFormatPatch{MarginLeft: formattingValue(int64(457200)), Indent: formattingValue(int64(-228600))}, token: []byte(`indent="-228600"`)},
		{name: "spacing", pp: ParagraphFormatPatch{SpaceBefore: formattingValue(int64(1200)), SpaceAfter: formattingValue(int64(600))}, token: []byte(`val="1200"`)},
		{name: "line", pp: ParagraphFormatPatch{LinePercent: formattingValue(int64(150000))}, token: []byte(`val="150000"`)},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			source, s := formattingFixture(t)
			original := bytes.Clone(source)
			target, e := s.FindFormatShape("ppt/slides/slide2.xml", 4)
			if e != nil {
				t.Fatal(e)
			}
			if tc.visibility != nil {
				e = s.SetSlideVisibility("ppt/slides/slide2.xml", *tc.visibility)
			} else if tc.geometry.X != nil || tc.geometry.Width != nil || tc.geometry.Rotation != nil {
				e = s.SetFormatGeometry(target, tc.geometry)
			} else if tc.rp.Bold != nil || tc.rp.Italic != nil || tc.rp.Latin != nil || tc.rp.Color != nil || tc.rp.RemoveDirect {
				e = s.SetFormatRun(target, 0, tc.run, tc.rp)
			} else {
				e = s.SetFormatParagraph(target, 0, tc.pp)
			}
			if e != nil {
				t.Fatal(e)
			}
			saved := formattingSave(t, s)
			if bytes.Equal(source, saved) {
				t.Fatal("no saved change")
			}
			if !bytes.Equal(source, original) {
				t.Fatal("caller source mutated")
			}
			reopened, e := OpenEditing(saved, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			b, _, e := reopened.pkg.Part("ppt/slides/slide2.xml")
			if e != nil || !bytes.Contains(b, tc.token) {
				t.Fatalf("missing reopened direct property %s: %v", tc.token, e)
			}
		})
	}
}
func TestFormattingAttributeLexicalBoundaries(t *testing.T) {
	original := []byte(`<a:rPr lang='en-US sz="1800" note' sz='1800' i="0"/>`)
	updated, e := fmtOpening(original, map[string]*string{"sz": formattingValue("2400")})
	if e != nil {
		t.Fatal(e)
	}
	if !bytes.Contains(updated, []byte(`lang='en-US sz="1800" note'`)) || !bytes.Contains(updated, []byte(`sz='2400'`)) || !bytes.Contains(updated, []byte(`i="0"`)) {
		t.Fatalf("quoted attribute bleed: %s", updated)
	}
	unchanged, e := fmtOpening(original, map[string]*string{"sz": formattingValue("1800")})
	if e != nil || !bytes.Equal(unchanged, original) {
		t.Fatalf("same value lost lexical form: %s %v", unchanged, e)
	}
	typeface := `Aptos <& "Quoted" 'Unicode Ω'`
	if _, e = fmtOpening([]byte(`<a:latin typeface='Arial'/>`), map[string]*string{"typeface": &typeface}); e != nil {
		t.Fatal(e)
	}
	for _, bad := range []string{`<a:rPr lang='broken sz="1800"/>`, `<a:rPr sz=1800/>`} {
		if _, e = fmtOpening([]byte(bad), map[string]*string{"sz": formattingValue("2400")}); e == nil {
			t.Fatalf("accepted malformed opening %s", bad)
		}
	}
}

// A quoted > belongs to an unpatched value, not to the end of the tag.
// Save/reopen proves the edit acts on the selected property and retains it.
func TestFormattingQuotedGreaterThanProfile(t *testing.T) {
	cases := []struct {
		name, old, quoted, changed string
		apply                      func(*EditSession, *FormatTarget) error
	}{
		{"run", `lang="en-US"`, `lang="en-US" note="a > b"`, `sz="2400"`, func(s *EditSession, target *FormatTarget) error {
			return s.SetFormatRun(target, 0, 0, RunFormatPatch{Size: formattingValue(int64(2400))})
		}},
		{"paragraph", `algn="l"`, `algn="l" note="a > b"`, `marL="457200"`, func(s *EditSession, target *FormatTarget) error {
			return s.SetFormatParagraph(target, 0, ParagraphFormatPatch{MarginLeft: formattingValue(int64(457200))})
		}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			_, s := formattingFixture(t)
			manipulationChangePart(t, s, "ppt/slides/slide2.xml", tc.old, tc.quoted)
			before := formattingSave(t, s)
			target, err := s.FindFormatShape("ppt/slides/slide2.xml", 4)
			if err != nil {
				t.Fatal(err)
			}
			if err := tc.apply(s, target); err != nil {
				t.Fatal(err)
			}
			saved := formattingSave(t, s)
			if bytes.Equal(before, saved) {
				t.Fatal("edit was a no-op")
			}
			reopened, err := OpenEditing(saved, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			slide, _, err := reopened.pkg.Part("ppt/slides/slide2.xml")
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Contains(slide, []byte(tc.quoted)) || !bytes.Contains(slide, []byte(tc.changed)) {
				t.Fatal("reopened selected property or unpatched quoted value lost")
			}
		})
	}
}

func TestFormattingAlphaAtomicRefusal(t *testing.T) {
	_, s := formattingFixture(t)
	manipulationChangePart(t, s, "ppt/slides/slide2.xml", `<a:srgbClr val="112233"/>`, `<a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr>`)
	before := formattingSave(t, s)
	target, err := s.FindFormatShape("ppt/slides/slide2.xml", 4)
	if err != nil {
		t.Fatal(err)
	}
	formattingRefusal(t, s.SetFormatRun(target, 0, 0, RunFormatPatch{Color: formattingValue("A1B2C3")}), "unsupported_structure")
	if err := s.formatCurrent(target); err != nil {
		t.Fatalf("refusal invalidated held target: %v", err)
	}
	if err := s.SetFormatParagraph(target, 0, ParagraphFormatPatch{Alignment: formattingValue("l")}); err != nil {
		t.Fatalf("held target unusable: %v", err)
	}
	if !bytes.Equal(before, formattingSave(t, s)) {
		t.Fatal("alpha refusal or follow-up no-op mutated session")
	}
}

func TestFormattingTopologyRefusals(t *testing.T) {
	cases := []struct{ name, part, old, new, kind string }{
		{"duplicate shape ID", "ppt/slides/slide2.xml", `<p:cNvPr id="4" name="Body"/>`, `<p:cNvPr id="2" name="Body"/>`, "ambiguous_target"},
		{"field run", "ppt/slides/slide2.xml", `<a:r><a:rPr lang="en-US"`, `<a:r><a:fld/><a:rPr lang="en-US"`, "unsupported_structure"},
		{"hyperlink", "ppt/slides/slide2.xml", `<a:rPr lang="en-US"`, `<a:rPr><a:hlinkClick r:id="rId9"/></a:rPr><a:rPr lang="en-US"`, "unsupported_structure"},
		{"locked shape", "ppt/slides/slide2.xml", `<p:cNvSpPr txBox="1"/>`, `<p:cNvSpPr txBox="1"><a:spLocks noSelect="1"/></p:cNvSpPr>`, "protected_operation"},
		{"unknown extension", "ppt/slides/slide2.xml", `<p:cNvSpPr txBox="1"/>`, `<p:cNvSpPr txBox="1"><a:extLst/></p:cNvSpPr>`, "unsupported_structure"},
		{"protected presentation", "ppt/presentation.xml", "</p:presentation>", "<p:modifyVerifier/></p:presentation>", "protected_operation"},
		{"duplicate run properties", "ppt/slides/slide2.xml", `<a:rPr lang="en-US"`, `<a:rPr/><a:rPr lang="en-US"`, "unsupported_structure"},
		{"foreign visibility", "ppt/slides/slide2.xml", `<p:sld xmlns:a=`, `<p:sld p:show="0" xmlns:a=`, "invalid-visibility"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			_, s := formattingFixture(t)
			manipulationChangePart(t, s, tc.part, tc.old, tc.new)
			before := formattingSave(t, s)
			var err error
			if tc.name == "foreign visibility" {
				err = s.SetSlideVisibility("ppt/slides/slide2.xml", false)
			} else {
				var target *FormatTarget
				target, err = s.FindFormatShape("ppt/slides/slide2.xml", 4)
				if err == nil {
					err = s.SetFormatRun(target, 0, 0, RunFormatPatch{Bold: formattingValue(true)})
				}
			}
			formattingRefusal(t, err, tc.kind)
			if !bytes.Equal(before, formattingSave(t, s)) {
				t.Fatal("refusal changed archive")
			}
		})
	}
	_, s := formattingFixture(t)
	first, e := s.FindFormatShape("ppt/slides/slide2.xml", 4)
	if e != nil {
		t.Fatal(e)
	}
	other := *first
	other.session = nil
	formattingRefusal(t, s.SetFormatRun(&other, 0, 0, RunFormatPatch{Bold: formattingValue(true)}), "stale_target")
}

func TestFormattingAtomicRefusals(t *testing.T) {
	_, s := formattingFixture(t)
	held, e := s.FindFormatShape("ppt/slides/slide2.xml", 4)
	if e != nil {
		t.Fatal(e)
	}
	before := formattingSave(t, s)
	checks := []struct {
		name, kind string
		run        func() error
	}{
		{"invalid visibility", "invalid-visibility", func() error { return s.SetSlideVisibilityValue("ppt/slides/slide2.xml", "hidden") }},
		{"late geometry", "invalid-geometry", func() error {
			return s.SetFormatGeometry(held, GeometryPatch{X: formattingValue(int64(914400)), Width: formattingValue(int64(-1))})
		}},
		{"bad rotation", "invalid-geometry", func() error {
			return s.SetFormatGeometry(held, GeometryPatch{Rotation: formattingValue(int64(21600000))})
		}},
		{"font bounds", "unsupported_structure", func() error { return s.SetFormatRun(held, 0, 0, RunFormatPatch{Size: formattingValue(int64(99))}) }},
		{"bad paragraph spacing", "unsupported_structure", func() error {
			return s.SetFormatParagraph(held, 0, ParagraphFormatPatch{LinePercent: formattingValue(int64(0))})
		}},
		{"foreign target", "stale_target", func() error {
			_, other := formattingFixture(t)
			return other.SetFormatRun(held, 0, 0, RunFormatPatch{Bold: formattingValue(true)})
		}},
	}
	for _, tc := range checks {
		t.Run(tc.name, func(t *testing.T) {
			formattingRefusal(t, tc.run(), tc.kind)
			if e = s.formatCurrent(held); e != nil {
				t.Fatalf("held target changed: %v", e)
			}
			if !bytes.Equal(before, formattingSave(t, s)) {
				t.Fatal("refusal mutated session")
			}
		})
	}
}
