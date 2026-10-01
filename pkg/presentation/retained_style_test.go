package presentation

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestRetainedStyleNative(t *testing.T) {
	path := testutil.ReferencePath("fixtures", "pptx", "shape-style", "shape-style-b4e7fd03880f.pptx")
	source, e := os.ReadFile(path)
	if os.IsNotExist(e) {
		t.Skip("sealed candidate absent")
	}
	if e != nil {
		t.Fatal(e)
	}
	fill := "A1B2C3"
	none := "none"
	color := "334455"
	dash := "dash"
	anchor := "ctr"
	strike := "sngStrike"
	width := int64(25400)
	columns := int64(2)
	gap := int64(91440)
	cases := []struct {
		name    string
		patch   RetainedShapeStylePatch
		body    RetainedBodyPatch
		run     RetainedRunPatch
		refusal bool
		token   []byte
	}{
		{name: "fill", patch: RetainedShapeStylePatch{Fill: &fill}, token: []byte(`val="A1B2C3"`)},
		{name: "noFill", patch: RetainedShapeStylePatch{Fill: &none}, token: []byte(`<a:noFill/>`)},
		{name: "inherit", patch: RetainedShapeStylePatch{InheritFill: true}, token: []byte(`<a:prstGeom`)},
		{name: "lineColor", patch: RetainedShapeStylePatch{LineColor: &color}, token: []byte(`val="334455"`)},
		{name: "lineWidth", patch: RetainedShapeStylePatch{LineWidth: &width}, token: []byte(`w="25400"`)},
		{name: "dash", patch: RetainedShapeStylePatch{LineDash: &dash}, token: []byte(`<a:prstDash val="dash"/>`)},
		{name: "body", body: RetainedBodyPatch{Anchor: &anchor}, token: []byte(`anchor="ctr"`)},
		{name: "columns", body: RetainedBodyPatch{NumCol: &columns, SpcCol: &gap}, token: []byte(`numCol="2"`)},
		{name: "run", run: RetainedRunPatch{Strike: &strike}, token: []byte(`strike="sngStrike"`)},
		{name: "refusal", patch: RetainedShapeStylePatch{LineColor: &color, LineWidth: func() *int64 { x := int64(-1); return &x }()}, refusal: true},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			original := bytes.Clone(source)
			s, e := OpenEditing(source, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			target, e := s.FindFormatShape("ppt/slides/slide2.xml", 4)
			if e != nil {
				t.Fatal(e)
			}
			switch {
			case c.name == "body" || c.name == "columns":
				e = s.SetRetainedBody(target, c.body)
			case c.name == "run":
				e = s.SetRetainedRun(target, 0, 0, c.run)
			default:
				e = s.SetRetainedShapeStyle(target, c.patch)
			}
			if c.refusal {
				if e == nil {
					t.Fatal("missing refusal")
				}
			} else if e != nil {
				t.Fatal(e)
			}
			out := filepath.Join(t.TempDir(), "out.pptx")
			if _, e = s.SaveAs(out); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(out)
			if e != nil {
				t.Fatal(e)
			}
			if c.refusal {
				if !bytes.Equal(saved, source) {
					t.Fatal("refusal mutated archive")
				}
				return
			}
			if bytes.Equal(saved, source) {
				t.Fatal("no edit")
			}
			if !bytes.Equal(source, original) {
				t.Fatal("actual caller buffer changed")
			}
			reopened, e := OpenEditing(saved, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			part, _, e := reopened.pkg.Part("ppt/slides/slide2.xml")
			if e != nil || !bytes.Contains(part, c.token) {
				t.Fatalf("missing reopened token %q: %v", c.token, e)
			}
		})
	}
}
