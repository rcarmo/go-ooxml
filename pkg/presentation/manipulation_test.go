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

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const manipulationFixtureSHA = "5ad4b68acc5a926c020dbc94154c95fa6c40c7353a55b2f82504879b392e9a45"

func manipulationFixture(t *testing.T) ([]byte, *EditSession) {
	t.Helper()
	path := testutil.ReferencePath("fixtures", "pptx", "manipulation", "manipulation-5ad4b68acc5a.pptx")
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skip("PPTX manipulation fixture is candidate-only")
	}
	source, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(source)
	if hex.EncodeToString(sum[:]) != manipulationFixtureSHA {
		t.Fatal("candidate fixture SHA differs")
	}
	s, err := OpenEditing(source, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	return source, s
}
func manipulationBytes(t *testing.T, s *EditSession) []byte {
	t.Helper()
	p := filepath.Join(t.TempDir(), "edited.pptx")
	if _, err := s.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	b, err := os.ReadFile(p)
	if err != nil {
		t.Fatal(err)
	}
	return b
}
func manipulationKind(t *testing.T, err error, kind string) {
	t.Helper()
	var refused *packaging.Refusal
	if !errors.As(err, &refused) || refused.Kind != kind {
		t.Fatalf("refusal=%v want %s", err, kind)
	}
}
func TestRetainedManipulationNativeControls(t *testing.T) {
	source, s := manipulationFixture(t)
	before := bytes.Clone(source)
	title, err := s.FindShape("ppt/slides/slide1.xml", 2)
	if err != nil || title.Text() != "Alpha" {
		t.Fatalf("title=%v %v", title, err)
	}
	if _, err = s.FindShape("ppt/slides/slide1.xml", 99); err == nil {
		t.Fatal("missing shape accepted")
	}
	if _, err = s.FindShape("ppt/slides/slide99.xml", 2); err == nil {
		t.Fatal("unenrolled slide accepted")
	}
	if err = s.SetShapeText(title, " ", false); err != nil {
		t.Fatal(err)
	}
	manipulationKind(t, s.SetShapeText(title, "wrong", false), "stale_target")
	held, err := s.FindShape("ppt/slides/slide1.xml", 2)
	if err != nil {
		t.Fatal(err)
	}
	noOp := manipulationBytes(t, s)
	manipulationKind(t, s.AddShapeBullet(held, "too deep", 9, ""), "unsupported_structure")
	manipulationKind(t, s.SetShapeAutofit(held, "unknown"), "unsupported_structure")
	manipulationKind(t, s.SetShapeText(held, "bad\r", false), "unsupported_structure")
	manipulationKind(t, s.ReorderSlides([]int{0, 0, 1}), "invalid-permutation")
	if held.Text() != " " || !bytes.Equal(noOp, manipulationBytes(t, s)) || !bytes.Equal(source, before) {
		t.Fatal("refusal changed session, held target, or caller bytes")
	}
	if err = s.SetShapeText(held, "\n A&B \n雪\n", false); err != nil {
		t.Fatal(err)
	}
	data := manipulationBytes(t, s)
	p, err := packaging.OpenPreserved(data, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	slide, _, err := p.Part("ppt/slides/slide1.xml")
	if err != nil {
		t.Fatal(err)
	}
	d, err := losslessxml.Parse(slide)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(slide, []byte("A&amp;B")) || len(d.Elements()) == 0 {
		t.Fatal("Unicode/ampersand/empty paragraph lost")
	}
	// A fault after staging a graph addition must refuse before the only apply.
	_, second := manipulationFixture(t)
	pre := manipulationBytes(t, second)
	g, e := second.pkg.Graph()
	if e != nil {
		t.Fatal(e)
	}
	if len(g.Edges) == 0 {
		t.Fatal("missing original graph")
	}
	plan, e := second.pkg.PlanGraphMutation(packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: "ppt/slides/slide4.xml", ContentType: packaging.ContentTypeSlide, Data: []byte(`<p:sld xmlns:p="` + packaging.NSPresentationML + `"/>`)}}, Relationships: []packaging.RelationshipAddition{{Source: "ppt/slides/slide4.xml", ID: "rId1", Type: packaging.RelTypeSlideLayout, TargetPart: "missing/layout.xml"}}})
	if e == nil || plan != nil {
		t.Fatalf("dangling pending-owner plan accepted: %v", e)
	}
	if !bytes.Equal(pre, manipulationBytes(t, second)) {
		t.Fatal("graph preflight refusal changed archive")
	}
	if err = second.InsertTextSlide(-1, "bad", "subtitle"); err == nil {
		t.Fatal("negative insert accepted")
	}
	if !bytes.Equal(pre, manipulationBytes(t, second)) {
		t.Fatal("insert refusal changed archive")
	}
	if err = second.TableValues("ppt/slides/slide1.xml", [][]string{{"x"}, {"x", "y"}}, 0, 0, 100, 100); err == nil {
		t.Fatal("ragged table accepted")
	}
	if err = second.SetShapeText(nil, "x", false); err == nil {
		t.Fatal("nil target accepted")
	}
}

func manipulationChangePart(t *testing.T, s *EditSession, part, old, replacement string) {
	t.Helper()
	body, hash, err := s.pkg.Part(part)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Count(body, []byte(old)) != 1 {
		t.Fatalf("expected exactly one %q in %s", old, part)
	}
	changed := bytes.Replace(body, []byte(old), []byte(replacement), 1)
	if err := s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: changed}}); err != nil {
		t.Fatal(err)
	}
}

func TestRetainedManipulationGuardedRefusals(t *testing.T) {
	cases := []struct {
		name, part, old, new, kind string
		table                      bool
	}{
		{"ambiguous placeholder", "ppt/slides/slide3.xml", "<p:ph type=\"subTitle\" idx=\"1\"/>", "<p:ph type=\"ctrTitle\" idx=\"1\"/>", "ambiguous_target", false},
		{"duplicate shape id", "ppt/slides/slide1.xml", "<p:cNvPr id=\"3\" name=\"Subtitle 2\"/>", "<p:cNvPr id=\"2\" name=\"Subtitle 2\"/>", "ambiguous_target", true},
		{"shape ID zero", "ppt/slides/slide1.xml", "<p:cNvPr id=\"2\" name=\"Title 1\"/>", "<p:cNvPr id=\"0\" name=\"Title 1\"/>", "unsupported_structure", true},
		{"shape ID exhausted", "ppt/slides/slide1.xml", "<p:cNvPr id=\"2\" name=\"Title 1\"/>", "<p:cNvPr id=\"4294967295\" name=\"Title 1\"/>", "unsupported_structure", true},
		{"protected table", "ppt/presentation.xml", "</p:presentation>", "<p:modifyVerifier/></p:presentation>", "protected_operation", true},
		{"slide-list custom child", "ppt/presentation.xml", "</p:sldIdLst>", "<p:unexpected/></p:sldIdLst>", "unsupported_structure", false},
		{"slide-list comment", "ppt/presentation.xml", "</p:sldIdLst>", "<!-- kept --></p:sldIdLst>", "unsupported_structure", false},
		{"slide-list whitespace", "ppt/presentation.xml", "</p:sldIdLst>", " </p:sldIdLst>", "unsupported_structure", false},
		{"slide-entry custom attr", "ppt/presentation.xml", "<p:sldId id=\"256\" r:id=\"rId6\"/>", "<p:sldId id=\"256\" r:id=\"rId6\" flag=\"1\"/>", "unsupported_structure", false},
		{"slide ID beyond ECMA", "ppt/presentation.xml", "<p:sldId id=\"258\" r:id=\"rId8\"/>", "<p:sldId id=\"2147483648\" r:id=\"rId8\"/>", "ambiguous_target", false},
		{"slide ID allocation exhausted", "ppt/presentation.xml", "<p:sldId id=\"258\" r:id=\"rId8\"/>", "<p:sldId id=\"2147483647\" r:id=\"rId8\"/>", "unsupported_structure", false},
		{"duplicate run property", "ppt/slides/slide3.xml", "<a:r><a:t>Gamma</a:t></a:r>", "<a:r><a:rPr/><a:rPr/><a:t>Gamma</a:t></a:r>", "unsupported_structure", false},
		{"invalid run nesting", "ppt/slides/slide3.xml", "<a:r><a:t>Gamma</a:t></a:r>", "<a:r><a:r><a:t>Gamma</a:t></a:r></a:r>", "unsupported_structure", false},
		{"invalid body order", "ppt/slides/slide3.xml", "<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Gamma", "<a:lstStyle/><a:bodyPr/><a:p><a:r><a:t>Gamma", "unsupported_structure", false},
		{"table shape tree owner", "ppt/slides/slide1.xml", "<p:cSld><p:spTree>", "<p:cSld><p:other><p:spTree/></p:other><p:spTree>", "unsupported_structure", true},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			_, s := manipulationFixture(t)
			manipulationChangePart(t, s, tc.part, tc.old, tc.new)
			prior := manipulationBytes(t, s)
			var err error
			if tc.table {
				err = s.TableValues("ppt/slides/slide1.xml", [][]string{{"A"}}, 1, 1, 100, 100)
			} else {
				err = s.InsertTextSlide(0, "Title", "Subtitle")
			}
			manipulationKind(t, err, tc.kind)
			if !bytes.Equal(prior, manipulationBytes(t, s)) {
				t.Fatal("refusal changed archive")
			}
		})
	}
	_, s := manipulationFixture(t)
	before := manipulationBytes(t, s)
	for _, rows := range [][][]string{make([][]string, 65), {{strings.Repeat("X", 4097)}}, {{"x"}}} {
		var err error
		if len(rows) == 1 && rows[0][0] == "x" {
			err = s.TableValues("ppt/slides/slide1.xml", rows, 1, 1, 100000001, 1)
		} else {
			err = s.TableValues("ppt/slides/slide1.xml", rows, 1, 1, 100, 100)
		}
		manipulationKind(t, err, "unsupported_structure")
	}
	if !bytes.Equal(before, manipulationBytes(t, s)) {
		t.Fatal("budget refusals changed archive")
	}
}

func TestRetainedManipulationBulletAutofitRefusals(t *testing.T) {
	for _, tc := range []struct{ name, part, old, replacement, operation string }{
		{"bullet text newline", "", "", "", "text"},
		{"bullet label newline", "", "", "", "label"},
		{"ambiguous autofit", "ppt/slides/slide1.xml", "<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alpha", "<a:bodyPr><a:noAutofit/><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>Alpha", "autofit"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			_, s := manipulationFixture(t)
			if tc.part != "" {
				manipulationChangePart(t, s, tc.part, tc.old, tc.replacement)
			}
			before := manipulationBytes(t, s)
			shape, err := s.FindShape("ppt/slides/slide1.xml", 2)
			if err != nil {
				t.Fatal(err)
			}
			switch tc.operation {
			case "text":
				err = s.AddShapeBullet(shape, "first\nsecond", 0, "")
			case "label":
				err = s.AddShapeBullet(shape, "description", 0, "first\nsecond")
			default:
				err = s.SetShapeAutofit(shape, "shrink")
			}
			manipulationKind(t, err, "unsupported_structure")
			if shape.Text() != "Alpha" || !bytes.Equal(before, manipulationBytes(t, s)) {
				t.Fatal("refusal changed shape handle or archive")
			}
		})
	}
}

func TestRetainedManipulationPlaceholderIdentity(t *testing.T) {
	_, s := manipulationFixture(t)
	// Template text is not an identity: independently alter the title to an
	// unrelated value, then insert with the same two placeholder types.
	shape, err := s.FindShape("ppt/slides/slide3.xml", 2)
	if err != nil {
		t.Fatal(err)
	}
	if err = s.SetShapeText(shape, "Different template", false); err != nil {
		t.Fatal(err)
	}
	if err = s.InsertTextSlide(0, "First Slide", "Inserted subtitle"); err != nil {
		t.Fatal(err)
	}
	data := manipulationBytes(t, s)
	p, err := OpenEditing(data, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	inserted, err := p.FindShape("ppt/slides/slide4.xml", 2)
	if err != nil || inserted.Text() != "First Slide" {
		t.Fatalf("inserted title=%v %v", inserted, err)
	}
	original, err := p.FindShape("ppt/slides/slide3.xml", 2)
	if err != nil || original.Text() != "Different template" {
		t.Fatalf("original title=%v %v", original, err)
	}

	// A different existing subtitle and even a self-closing title leaf are
	// valid: the new slide receives explicit paragraphs from placeholders.
	_, empty := manipulationFixture(t)
	body, h, err := empty.pkg.Part("ppt/slides/slide3.xml")
	if err != nil {
		t.Fatal(err)
	}
	doc, err := losslessxml.Parse(body)
	if err != nil {
		t.Fatal(err)
	}
	var title, subtitle losslessxml.Element
	for _, e := range doc.Elements() {
		if e.Name() == name(packaging.NSDrawingML, "t") {
			text, _ := e.Text()
			if text == "Gamma" {
				title = e
			}
			if text == "Subtitle" {
				subtitle = e
			}
		}
	}
	if title.Ordinal() < 0 || subtitle.Ordinal() < 0 {
		t.Fatal("template text missing")
	}
	changed, err := doc.ReplaceText([]losslessxml.TextEdit{{Target: title, Text: ""}, {Target: subtitle, Text: "Other subtitle"}})
	if err != nil {
		t.Fatal(err)
	}
	if err = empty.pkg.Replace([]packaging.Replacement{{Part: "ppt/slides/slide3.xml", ExpectedSHA256: h, Data: changed}}); err != nil {
		t.Fatal(err)
	}
	if err = empty.InsertTextSlide(1, "Middle Slide", "Inserted subtitle"); err != nil {
		t.Fatal(err)
	}
	check := manipulationBytes(t, empty)
	reopened, err := OpenEditing(check, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	inserted, err = reopened.FindShape("ppt/slides/slide4.xml", 2)
	if err != nil || inserted.Text() != "Middle Slide" {
		t.Fatalf("empty-title template: %v %v", inserted, err)
	}
}
