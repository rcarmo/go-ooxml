package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type pptxCaseState struct {
	record     pptxManipulationRecord
	source     []byte
	sourceCopy []byte
	before     map[string][]byte
	session    *presentation.EditSession
	shape      *presentation.ShapeTarget
	notes      *presentation.NotesTarget
	held       *presentation.ShapeTarget
	refusal    error
	output     []byte
	after      map[string][]byte
}

const pptxP = "http://schemas.openxmlformats.org/presentationml/2006/main"
const pptxA = "http://schemas.openxmlformats.org/drawingml/2006/main"
const pptxR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

func pptxElem(ns, local string) xml.Name { return xml.Name{Space: ns, Local: local} }
func pptxChild(d *losslessxml.Document, p losslessxml.Element, n xml.Name) []losslessxml.Element {
	var out []losslessxml.Element
	for _, e := range d.Elements() {
		if parent, ok := e.Parent(); ok && parent == p && e.Name() == n {
			out = append(out, e)
		}
	}
	return out
}
func pptxDescendant(e, ancestor losslessxml.Element) bool {
	for p, ok := e.Parent(); ok; p, ok = p.Parent() {
		if p == ancestor {
			return true
		}
	}
	return false
}
func pptxAttr(e losslessxml.Element, ns, local string) string {
	for _, a := range e.Attributes() {
		if a.Name == pptxElem(ns, local) {
			return a.Value
		}
	}
	return ""
}
func pptxMembers(data []byte) (map[string][]byte, error) {
	z, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	parts := map[string][]byte{}
	for _, f := range z.File {
		if _, exists := parts[f.Name]; exists {
			return nil, fmt.Errorf("duplicate ZIP member %q", f.Name)
		}
		r, e := f.Open()
		if e != nil {
			return nil, e
		}
		b, e := io.ReadAll(r)
		closeErr := r.Close()
		if e != nil {
			return nil, e
		}
		if closeErr != nil {
			return nil, closeErr
		}
		parts[f.Name] = b
	}
	return parts, nil
}
func pptxShape(d *losslessxml.Document, id uint32) (losslessxml.Element, error) {
	var shape losslessxml.Element
	found := 0
	for _, e := range d.Elements() {
		if e.Name() != pptxElem(pptxP, "cNvPr") {
			continue
		}
		n, err := strconv.ParseUint(pptxAttr(e, "", "id"), 10, 32)
		if err != nil {
			return shape, err
		}
		if uint32(n) != id {
			continue
		}
		found++
		nv, ok := e.Parent()
		if !ok {
			return shape, fmt.Errorf("shape owner absent")
		}
		sp, ok := nv.Parent()
		if !ok || sp.Name() != pptxElem(pptxP, "sp") {
			return shape, fmt.Errorf("shape owner not plain")
		}
		shape = sp
	}
	if found != 1 {
		return shape, fmt.Errorf("shape ID %d has %d owners", id, found)
	}
	return shape, nil
}
func pptxTexts(d *losslessxml.Document, ancestor losslessxml.Element) []string {
	var texts []string
	for _, e := range d.Elements() {
		if e.Name() == pptxElem(pptxA, "t") && pptxDescendant(e, ancestor) {
			v, _ := e.Text()
			texts = append(texts, v)
		}
	}
	return texts
}
func pptxShapeParagraphs(d *losslessxml.Document, shape losslessxml.Element) ([]losslessxml.Element, []string, error) {
	body := pptxChild(d, shape, pptxElem(pptxP, "txBody"))
	if len(body) != 1 {
		return nil, nil, fmt.Errorf("shape body count %d", len(body))
	}
	ps := pptxChild(d, body[0], pptxElem(pptxA, "p"))
	lines := make([]string, 0, len(ps))
	for _, p := range ps {
		lines = append(lines, strings.Join(pptxTexts(d, p), ""))
	}
	return ps, lines, nil
}
func pptxSlideParts(archive []byte) ([]string, []string, []string, error) {
	gPkg, err := packaging.OpenPreserved(archive, packaging.Limits{})
	if err != nil {
		return nil, nil, nil, err
	}
	g, err := gPkg.Graph()
	if err != nil {
		return nil, nil, nil, err
	}
	main := ""
	for _, e := range g.Edges {
		if e.Source == "" && e.Type == packaging.RelTypeOfficeDocument && !e.External {
			if main != "" {
				return nil, nil, nil, fmt.Errorf("multiple presentation roots")
			}
			main = e.ResolvedPart
		}
	}
	part, _, err := gPkg.Part(main)
	if err != nil {
		return nil, nil, nil, err
	}
	d, err := losslessxml.Parse(part)
	if err != nil {
		return nil, nil, nil, err
	}
	var list losslessxml.Element
	for _, e := range d.Elements() {
		if e.Name() == pptxElem(pptxP, "sldIdLst") {
			list = e
			break
		}
	}
	rels := map[string]string{}
	for _, e := range g.Edges {
		if e.Source == main && e.Type == packaging.RelTypeSlide && !e.External {
			rels[e.ID] = e.ResolvedPart
		}
	}
	var paths, rids, ids []string
	for _, e := range pptxChild(d, list, pptxElem(pptxP, "sldId")) {
		rid := pptxAttr(e, pptxR, "id")
		p := rels[rid]
		if p == "" {
			return nil, nil, nil, fmt.Errorf("unresolved slide %s", rid)
		}
		paths = append(paths, p)
		rids = append(rids, rid)
		ids = append(ids, pptxAttr(e, "", "id"))
	}
	return paths, rids, ids, nil
}

func pptxJSON(raw json.RawMessage) (map[string]any, error) {
	var out map[string]any
	if err := json.Unmarshal(raw, &out); err != nil {
		return nil, err
	}
	return out, nil
}
func pptxList(value any) ([]string, error) {
	items, ok := value.([]any)
	if !ok {
		return nil, fmt.Errorf("not a string list: %T", value)
	}
	out := make([]string, len(items))
	for i, v := range items {
		out[i], ok = v.(string)
		if !ok {
			return nil, fmt.Errorf("non-string list value")
		}
	}
	return out, nil
}
func pptxSameStrings(got []string, expected any) error {
	want, err := pptxList(expected)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(got, want) {
		return fmt.Errorf("values=%q want=%q", got, want)
	}
	return nil
}
func pptxManipulationSteps(sc *godog.ScenarioContext) {
	var s pptxCaseState
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		s = pptxCaseState{}
		return ctx, nil
	})
	sc.Step(`^the sealed PPTX fixture (fixture-[0-9a-f]+) with ordered titles (\[.*\])$`, func(fixture, titles string) error {
		if fixture != pptxManipulationFixture {
			return fmt.Errorf("unreviewed fixture %q", fixture)
		}
		path := testutil.ReferencePath(filepath.FromSlash(pptxManipulationFixtureFile))
		b, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		sum := sha256.Sum256(b)
		if len(b) != 8464 || hex.EncodeToString(sum[:]) != strings.TrimPrefix(fixture, "fixture-") {
			return fmt.Errorf("fixture seal differs")
		}
		before, err := pptxMembers(b)
		if err != nil {
			return err
		}
		if len(before) != 20 {
			return fmt.Errorf("fixture member count %d", len(before))
		}
		var ledger struct {
			Fixture struct {
				Members []struct {
					Name   string `json:"name"`
					SHA256 string `json:"sha256"`
				} `json:"members"`
			} `json:"fixture"`
		}
		ledgerBytes, err := os.ReadFile(testutil.ReferencePath("ledgers", "pptx-manipulation.json"))
		if err != nil {
			return err
		}
		if err := json.Unmarshal(ledgerBytes, &ledger); err != nil {
			return err
		}
		if len(ledger.Fixture.Members) != len(before) {
			return fmt.Errorf("fixture ledger member count differs")
		}
		for _, member := range ledger.Fixture.Members {
			payload, ok := before[member.Name]
			if !ok {
				return fmt.Errorf("fixture member missing %s", member.Name)
			}
			seal := sha256.Sum256(payload)
			if hex.EncodeToString(seal[:]) != member.SHA256 {
				return fmt.Errorf("fixture member seal differs %s", member.Name)
			}
		}
		session, err := presentation.OpenEditing(b, packaging.Limits{})
		if err != nil {
			return err
		}
		s = pptxCaseState{source: b, sourceCopy: bytes.Clone(b), before: before, session: session}
		var want []string
		if err = json.Unmarshal([]byte(titles), &want); err != nil {
			return err
		}
		if !reflect.DeepEqual(want, []string{"Alpha", "Beta", "Gamma"}) {
			return fmt.Errorf("unexpected fixture titles")
		}
		return nil
	})
	sc.Step(`^slide (\d+) shape ID (\d+) is selected by its original part and unique identity$`, func(slide, id int) error {
		part := fmt.Sprintf("ppt/slides/slide%d.xml", slide)
		target, err := s.session.FindShape(part, uint32(id))
		if err != nil {
			return err
		}
		s.shape = target
		s.held = target
		return nil
	})
	sc.Step(`^slide (\d+) existing notes body is selected by its original part and unique identity$`, func(slide int) error {
		var err error
		s.notes, err = s.session.FindNotes(fmt.Sprintf("ppt/slides/slide%d.xml", slide))
		return err
	})
	sc.Step(`^slide (\d+) is selected by its original part and unique identity$`, func(slide int) error {
		if slide != 1 {
			return fmt.Errorf("unreviewed slide %d", slide)
		}
		return nil
	})
	sc.Step(`^the presentation slide list is selected by its original part and unique identity$`, func() error { return nil })
	sc.Step(`^production presentation editing APIs (.+)$`, func(operation string) error {
		records, err := pptxManipulationRecords()
		if err != nil {
			return err
		}
		matches := []pptxManipulationRecord{}
		for _, r := range records {
			if r.Operation == operation {
				matches = append(matches, r)
			}
		}
		if len(matches) != 1 {
			return fmt.Errorf("unreviewed manipulation %q", operation)
		}
		s.record = matches[0]
		kind := s.record.Kind
		var runErr error
		switch kind {
		case "patch-title", "patch-body", "patch-subtitle":
			v := map[string]string{"patch-title": "New Title", "patch-body": "Updated bullet", "patch-subtitle": "New Subtitle"}[kind]
			runErr = s.session.SetShapeText(s.shape, v, false)
		case "append-title":
			runErr = s.session.SetShapeText(s.shape, " - Appended", true)
		case "bullet-default":
			runErr = s.session.AddShapeBullet(s.shape, "New bullet point", 0, "")
		case "bullet-sequence":
			if runErr = s.session.ClearShapeText(s.shape); runErr == nil {
				s.shape, runErr = s.session.FindShape("ppt/slides/slide2.xml", 4)
			}
			if runErr == nil {
				runErr = s.session.AddShapeBullet(s.shape, "First bullet", 0, "")
			}
			if runErr == nil {
				s.shape, runErr = s.session.FindShape("ppt/slides/slide2.xml", 4)
			}
			if runErr == nil {
				runErr = s.session.AddShapeBullet(s.shape, "Second bullet", 0, "")
			}
		case "bullet-level":
			runErr = s.session.AddShapeBullet(s.shape, "Sub-bullet point", 1, "")
		case "bullet-bold-label":
			runErr = s.session.AddShapeBullet(s.shape, "This is the description", 0, "Key Point")
		case "clear-bullets":
			runErr = s.session.ClearShapeText(s.shape)
		case "autofit-shrink", "autofit-none", "autofit-resize":
			runErr = s.session.SetShapeAutofit(s.shape, strings.TrimPrefix(kind, "autofit-"))
		case "insert-start":
			runErr = s.session.InsertTextSlide(0, "First Slide", "Inserted subtitle")
		case "insert-middle":
			runErr = s.session.InsertTextSlide(1, "Middle Slide", "Inserted subtitle")
		case "reorder":
			runErr = s.session.ReorderSlides([]int{0, 2, 1})
		case "reorder-refusal":
			s.held, runErr = s.session.FindShape("ppt/slides/slide1.xml", 2)
			if runErr == nil {
				runErr = s.session.ReorderSlides([]int{0, 1, 2, 3, 4, 5})
			}
			s.refusal = runErr
			return nil
		case "table-values":
			runErr = s.session.TableValues("ppt/slides/slide1.xml", [][]string{{"Name", "Value", "Status"}, {"Item 1", "100", "OK"}, {"Item 2", "200", "Pending"}}, 914400, 1828800, 10058400, 2743200)
		case "table-geometry":
			runErr = s.session.TableValues("ppt/slides/slide1.xml", [][]string{{"", ""}, {"", ""}}, 914400, 1828800, 10058400, 2743200)
		case "set-notes":
			runErr = s.session.ReplaceNotes(s.notes, "These are speaker notes")
		case "notes-readback":
			runErr = s.session.ReplaceNotes(s.notes, "Speaker notes content\nsecond line")
		default:
			return fmt.Errorf("unreviewed kind %s", kind)
		}
		return runErr
	})
	sc.Step(`^the resulting presentation is saved to a new path and independently parsed and reopened$`, func() error {
		if s.record.Kind == "reorder-refusal" {
			return nil
		}
		dir, err := os.MkdirTemp("", "go-pptx-manipulation-")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		path := filepath.Join(dir, "result.pptx")
		if _, err = s.session.SaveAs(path); err != nil {
			return err
		}
		b, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		parts, err := pptxMembers(b)
		if err != nil {
			return err
		}
		p, err := packaging.OpenPreserved(b, packaging.Limits{})
		if err != nil {
			return err
		}
		if _, err = p.Graph(); err != nil {
			return err
		}
		s.output = b
		s.after = parts
		return nil
	})
	sc.Step(`^the saved (.+) has exact properties JSON (\{.*\})$`, func(target, expected string) error {
		if s.record.Target != target {
			return fmt.Errorf("target changed %q vs %q", target, s.record.Target)
		}
		var scenarioValue, ledgerValue any
		if err := json.Unmarshal([]byte(expected), &scenarioValue); err != nil {
			return err
		}
		if err := json.Unmarshal(s.record.Expected, &ledgerValue); err != nil {
			return err
		}
		if !reflect.DeepEqual(scenarioValue, ledgerValue) {
			return fmt.Errorf("scenario expected JSON differs from sealed ledger")
		}
		return s.validateExpected()
	})
	sc.Step(`^slide-order refusal preserves archive bytes and handle order$`, func() error {
		var refused *packaging.Refusal
		if !errors.As(s.refusal, &refused) || refused.Kind != "invalid-permutation" {
			return fmt.Errorf("expected invalid-permutation, got %v", s.refusal)
		}
		if s.held == nil || s.held.Text() != "Alpha" {
			return fmt.Errorf("held target changed")
		}
		p, err := s.session.FindShape("ppt/slides/slide1.xml", 2)
		if err != nil || p.Text() != "Alpha" {
			return fmt.Errorf("refusal invalidated held target: %v", err)
		}
		return nil
	})
	sc.Step(`^only original member payloads (\[.*\]) may change$`, func(encoded string) error {
		var listed []string
		if err := json.Unmarshal([]byte(encoded), &listed); err != nil {
			return err
		}
		if !reflect.DeepEqual(listed, s.record.ChangedMembers) {
			return fmt.Errorf("changed-member allowlist drift")
		}
		return s.validateCustody()
	})
	sc.Step(`^the original member name set is unchanged$`, func() error {
		if s.record.Kind == "reorder-refusal" {
			return s.validateCustody()
		}
		if len(s.after) != len(s.before) {
			return fmt.Errorf("member count changed")
		}
		for n := range s.before {
			if _, ok := s.after[n]; !ok {
				return fmt.Errorf("member %s missing", n)
			}
		}
		return nil
	})
	sc.Step(`^the supplied archive, unrelated member payloads and unaffected lexical XML spans retain their original bytes$`, func() error { return s.validateCustody() })
	sc.Step(`^the supplied archive, unrelated member payloads and unaffected lexical XML spans retain their original bytes, and the refusal returns invalid-permutation without changing session or handles$`, func() error {
		if err := s.validateCustody(); err != nil {
			return err
		}
		var refused *packaging.Refusal
		if !errors.As(s.refusal, &refused) || refused.Kind != "invalid-permutation" {
			return fmt.Errorf("wrong refusal: %v", s.refusal)
		}
		return nil
	})
	sc.Step(`^slide identities and unrelated package bytes are retained$`, func() error { return s.validateCustody() })
	sc.Step(`^all pre-existing slides retain their original IDs, relationship IDs, targets, layout links and notes links$`, func() error { return s.validateCustody() })
	sc.Step(`^exactly one slide XML part and its one layout relationship part are added with valid unique IDs and a resolved original layout; no members are removed$`, func() error { return s.validateCustody() })
}
