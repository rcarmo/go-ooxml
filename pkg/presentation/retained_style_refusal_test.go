package presentation

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func retainedStyleFixture(t *testing.T) ([]byte, *EditSession) {
	t.Helper()
	data, e := os.ReadFile(testutil.ReferencePath("fixtures", "pptx", "shape-style", "shape-style-b4e7fd03880f.pptx"))
	if os.IsNotExist(e) {
		t.Skip("candidate absent")
	}
	if e != nil {
		t.Fatal(e)
	}
	s, e := OpenEditing(data, packaging.Limits{})
	if e != nil {
		t.Fatal(e)
	}
	return data, s
}
func retainedStyleSaved(t *testing.T, s *EditSession) []byte {
	t.Helper()
	dest := filepath.Join(t.TempDir(), "out.pptx")
	if _, e := s.SaveAs(dest); e != nil {
		t.Fatal(e)
	}
	b, e := os.ReadFile(dest)
	if e != nil {
		t.Fatal(e)
	}
	return b
}
func retainedStyleRefusal(t *testing.T, e error, kind string) {
	t.Helper()
	var r *packaging.Refusal
	if !errors.As(e, &r) || r.Kind != kind {
		t.Fatalf("refusal=%v want %s", e, kind)
	}
}
func TestRetainedStyleAtomicAndTopology(t *testing.T) {
	for _, c := range []struct {
		name  string
		patch RetainedShapeStylePatch
		body  RetainedBodyPatch
		kind  string
	}{
		{"late width", RetainedShapeStylePatch{LineColor: func() *string { x := "334455"; return &x }(), LineWidth: func() *int64 { x := int64(-1); return &x }()}, RetainedBodyPatch{}, "invalid-style"},
		{"body columns", RetainedShapeStylePatch{}, RetainedBodyPatch{Anchor: func() *string { x := "ctr"; return &x }(), NumCol: func() *int64 { x := int64(17); return &x }()}, "invalid-style"},
	} {
		t.Run(c.name, func(t *testing.T) {
			_, s := retainedStyleFixture(t)
			held, e := s.FindFormatShape("ppt/slides/slide2.xml", 4)
			if e != nil {
				t.Fatal(e)
			}
			before := retainedStyleSaved(t, s)
			if c.name == "body columns" {
				e = s.SetRetainedBody(held, c.body)
			} else {
				e = s.SetRetainedShapeStyle(held, c.patch)
			}
			retainedStyleRefusal(t, e, c.kind)
			if e = s.formatCurrent(held); e != nil {
				t.Fatalf("held target changed: %v", e)
			}
			same := "445566"
			if e = s.SetRetainedShapeStyle(held, RetainedShapeStylePatch{LineColor: &same}); e != nil {
				t.Fatalf("held no-op unusable: %v", e)
			}
			if !bytes.Equal(before, retainedStyleSaved(t, s)) {
				t.Fatal("refusal changed archive")
			}
		})
	}
	for _, c := range []struct{ name, old, new string }{
		{"duplicate fill", `<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln>`, `<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:noFill/><a:ln>`},
		{"alpha", `<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln>`, `<a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:ln>`},
		{"quoted style", `<a:prstGeom prst="rect">`, `<a:prstGeom prst="rect" note="a > b">`},
	} {
		t.Run(c.name, func(t *testing.T) {
			_, s := retainedStyleFixture(t)
			manipulationChangePart(t, s, "ppt/slides/slide2.xml", c.old, c.new)
			before := retainedStyleSaved(t, s)
			held, e := s.FindFormatShape("ppt/slides/slide2.xml", 4)
			if e != nil {
				t.Fatal(e)
			}
			fill := "A1B2C3"
			e = s.SetRetainedShapeStyle(held, RetainedShapeStylePatch{Fill: &fill})
			if c.name == "quoted style" {
				if e != nil {
					t.Fatal(e)
				}
				return
			}
			retainedStyleRefusal(t, e, "unsupported_structure")
			if !bytes.Equal(before, retainedStyleSaved(t, s)) {
				t.Fatal("topology refusal mutated archive")
			}
		})
	}
}
