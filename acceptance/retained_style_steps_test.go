package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type retainedState struct {
	record                  retainedRecord
	caller, original, saved []byte
	before, after           map[string][]byte
	pptx                    *presentation.EditSession
	word                    *document.EditSession
	shape                   *presentation.FormatTarget
	target                  *document.RetainedFormatTarget
	refusal                 error
}

func retainedJSON(raw string, want json.RawMessage) error {
	var a, b any
	if e := json.Unmarshal([]byte(raw), &a); e != nil {
		return e
	}
	if e := json.Unmarshal(want, &b); e != nil {
		return e
	}
	if !reflect.DeepEqual(a, b) {
		return fmt.Errorf("sealed JSON mismatch: %s want %s", raw, want)
	}
	return nil
}
func retainedPatch(raw json.RawMessage) (map[string]any, error) {
	var x map[string]any
	if e := json.Unmarshal(raw, &x); e != nil {
		return nil, e
	}
	if x == nil {
		return nil, fmt.Errorf("empty patch")
	}
	return x, nil
}
func retainedInt(x any) (int64, error) {
	n, ok := x.(float64)
	if !ok || float64(int64(n)) != n {
		return 0, fmt.Errorf("integer operand %v", x)
	}
	return int64(n), nil
}
func retainedString(x any) (string, error) {
	s, ok := x.(string)
	if !ok {
		return "", fmt.Errorf("string operand %v", x)
	}
	return s, nil
}
func retainedBool(x any) (bool, error) {
	b, ok := x.(bool)
	if !ok {
		return false, fmt.Errorf("boolean operand %v", x)
	}
	return b, nil
}
func (s *retainedState) perform() error {
	patch, e := retainedPatch(s.record.Patch)
	if e != nil {
		return e
	}
	if s.record.Format == "docx" {
		converted := map[string]any{}
		if v, ok := patch["font"]; ok {
			converted["ascii"] = v
			converted["hAnsi"] = v
			delete(patch, "font")
		}
		if v, ok := patch["remove"]; ok {
			converted["removeDirect"] = true
			delete(patch, "remove")
			_ = v
		}
		for k, v := range patch {
			if n, ok := v.(float64); ok {
				value, er := retainedInt(n)
				if er != nil {
					return er
				}
				converted[k] = value
			} else {
				converted[k] = v
			}
		}
		return s.word.SetRetainedFormat(s.target, converted)
	}
	var style presentation.RetainedShapeStylePatch
	var body presentation.RetainedBodyPatch
	var run presentation.RetainedRunPatch
	for k, v := range patch {
		switch k {
		case "fill":
			if v == nil {
				style.InheritFill = true
			} else {
				x, er := retainedString(v)
				if er != nil {
					return er
				}
				style.Fill = &x
			}
		case "lineColor":
			x, er := retainedString(v)
			if er != nil {
				return er
			}
			style.LineColor = &x
		case "lineWidth":
			x, er := retainedInt(v)
			if er != nil {
				return er
			}
			style.LineWidth = &x
		case "lineDash":
			x, er := retainedString(v)
			if er != nil {
				return er
			}
			style.LineDash = &x
		case "anchor":
			x, er := retainedString(v)
			if er != nil {
				return er
			}
			body.Anchor = &x
		case "vert":
			x, er := retainedString(v)
			if er != nil {
				return er
			}
			body.Vert = &x
		case "wrap":
			x, er := retainedString(v)
			if er != nil {
				return er
			}
			body.Wrap = &x
		case "anchorCtr":
			x, er := retainedBool(v)
			if er != nil {
				return er
			}
			body.AnchorCtr = &x
		case "rtlCol":
			x, er := retainedBool(v)
			if er != nil {
				return er
			}
			body.RTLCol = &x
		case "rot", "lIns", "tIns", "rIns", "bIns", "numCol", "spcCol":
			x, er := retainedInt(v)
			if er != nil {
				return er
			}
			switch k {
			case "rot":
				body.Rot = &x
			case "lIns":
				body.LIns = &x
			case "tIns":
				body.TIns = &x
			case "rIns":
				body.RIns = &x
			case "bIns":
				body.BIns = &x
			case "numCol":
				body.NumCol = &x
			case "spcCol":
				body.SpcCol = &x
			}
		case "strike", "cap":
			x, er := retainedString(v)
			if er != nil {
				return er
			}
			if k == "strike" {
				run.Strike = &x
			} else {
				run.Cap = &x
			}
		case "baseline", "spc":
			x, er := retainedInt(v)
			if er != nil {
				return er
			}
			if k == "baseline" {
				run.Baseline = &x
			} else {
				run.Spacing = &x
			}
		default:
			return fmt.Errorf("unbound retained patch key %s", k)
		}
	}
	switch {
	case strings.HasPrefix(s.record.Kind, "shape-") || strings.HasPrefix(s.record.Kind, "line-") || s.record.Kind == "style-refusal":
		return s.pptx.SetRetainedShapeStyle(s.shape, style)
	case strings.HasPrefix(s.record.Kind, "body-"):
		return s.pptx.SetRetainedBody(s.shape, body)
	case strings.HasPrefix(s.record.Kind, "run-"):
		return s.pptx.SetRetainedRun(s.shape, 0, 0, run)
	}
	return fmt.Errorf("unsupported retained operation %s", s.record.Kind)
}
func (s *retainedState) part() string {
	if s.record.Format == "pptx" {
		return "ppt/slides/slide2.xml"
	}
	return "word/document.xml"
}
func (s *retainedState) custody() error {
	if s.saved == nil || s.after == nil {
		return fmt.Errorf("saved archive absent")
	}
	if !bytes.Equal(s.caller, s.original) {
		return fmt.Errorf("caller buffer changed")
	}
	if len(s.before) != len(s.after) {
		return fmt.Errorf("archive member count changed")
	}
	expected := map[string]bool{}
	for _, name := range s.record.ChangedMembers {
		expected[name] = true
	}
	actual := map[string]bool{}
	for name, old := range s.before {
		now, ok := s.after[name]
		if !ok {
			return fmt.Errorf("archive member removed %s", name)
		}
		if !bytes.Equal(old, now) {
			actual[name] = true
		}
	}
	if !reflect.DeepEqual(actual, expected) {
		return fmt.Errorf("changed members %v want %v", actual, expected)
	}
	before, e := packaging.OpenPreserved(s.original, packaging.Limits{})
	if e != nil {
		return e
	}
	after, e := packaging.OpenPreserved(s.saved, packaging.Limits{})
	if e != nil {
		return e
	}
	a, e := before.Graph()
	if e != nil {
		return e
	}
	b, e := after.Graph()
	if e != nil {
		return e
	}
	if !reflect.DeepEqual(a, b) {
		return fmt.Errorf("package graph changed")
	}
	if s.record.Format == "pptx" {
		oldPaths, oldRids, oldIds, _, er := pptxSlideTitles(s.original)
		if er != nil {
			return er
		}
		paths, rids, ids, _, er := pptxSlideTitles(s.saved)
		if er != nil {
			return er
		}
		if !reflect.DeepEqual(oldPaths, paths) || !reflect.DeepEqual(oldRids, rids) || !reflect.DeepEqual(oldIds, ids) {
			return fmt.Errorf("slide identity drift")
		}
	}
	return retainedMaskedPartSame(s.before[s.part()], s.after[s.part()], s.record)
}
func retainedSteps(sc *godog.ScenarioContext) {
	state := &retainedState{}
	sc.Before(func(ctx context.Context, scenario *godog.Scenario) (context.Context, error) {
		*state = retainedState{}
		records, e := retainedRecords()
		if e != nil {
			return ctx, e
		}
		for _, tag := range scenario.Tags {
			if strings.HasPrefix(tag.Name, "@id-pptx-retained-") || strings.HasPrefix(tag.Name, "@id-docx-retained-") {
				state.record = records[tag.Name]
			}
		}
		return ctx, nil
	})
	sc.Step(`^retained (pptx|docx) input (fixture-[0-9a-f]+) has its sealed original property and identity snapshot$`, func(format, id string) error {
		if state.record.ID == "" || format != state.record.Format || id != state.record.FixtureID {
			return fmt.Errorf("retained fixture record mismatch")
		}
		fixture := map[string]string{"pptx": "fixtures/pptx/shape-style/shape-style-b4e7fd03880f.pptx", "docx": "fixtures/docx/direct-formatting/direct-formatting-2d32cedb722e.docx"}[format]
		b, e := os.ReadFile(testutil.ReferencePath(filepath.FromSlash(fixture)))
		if e != nil {
			return e
		}
		hash := sha256.Sum256(b)
		if "fixture-"+hex.EncodeToString(hash[:]) != id {
			return fmt.Errorf("retained source provenance")
		}
		state.caller = b
		state.original = bytes.Clone(b)
		state.before, e = pptxMembers(b)
		if e != nil {
			return e
		}
		if format == "pptx" {
			state.pptx, e = presentation.OpenEditing(b, packaging.Limits{})
		} else {
			state.word, e = document.OpenEditing(b, packaging.Limits{})
		}
		return e
	})
	sc.Step(`^retained edit target (.+) is uniquely selected with a held identity$`, func(target string) error {
		if target != state.record.Target {
			return fmt.Errorf("retained target %s", target)
		}
		var e error
		if state.pptx != nil {
			state.shape, e = state.pptx.FindFormatShape("ppt/slides/slide2.xml", 4)
		} else {
			run := -1
			if strings.HasPrefix(state.record.Kind, "run-") {
				run = 0
			}
			state.target, e = state.word.FindRetainedFormat(0, run)
		}
		return e
	})
	sc.Step(`^production retained (pptx|docx) editing applies ([a-z-]+) patch JSON (.+)$`, func(format, kind, raw string) error {
		if format != state.record.Format || kind != state.record.Kind {
			return fmt.Errorf("retained operation mismatch")
		}
		if e := retainedJSON(raw, state.record.Patch); e != nil {
			return e
		}
		e := state.perform()
		if state.record.Kind == "style-refusal" {
			state.refusal = e
			return nil
		}
		return e
	})
	sc.Step(`^the result or unchanged refusal session is saved and independently parsed and reopened$`, func() error {
		dest := filepath.Join(os.TempDir(), "go-retained-"+strings.TrimPrefix(state.record.ID, "@id-")+"-output."+state.record.Format)
		defer os.Remove(dest)
		var e error
		if state.pptx != nil {
			_, e = state.pptx.SaveAs(dest)
		} else {
			_, e = state.word.SaveAs(dest)
		}
		if e != nil {
			return e
		}
		state.saved, e = os.ReadFile(dest)
		if e != nil {
			return e
		}
		state.after, e = pptxMembers(state.saved)
		if e != nil {
			return e
		}
		if state.pptx != nil {
			_, e = presentation.OpenEditing(state.saved, packaging.Limits{})
		} else {
			_, e = document.OpenEditing(state.saved, packaging.Limits{})
		}
		return e
	})
	sc.Step(`^retained target (.+) has exact saved properties JSON (.+)$`, func(target, raw string) error {
		if target != state.record.Target {
			return fmt.Errorf("retained target mismatch")
		}
		if e := retainedJSON(raw, state.record.Expected); e != nil {
			return e
		}
		return retainedExpected(state)
	})
	sc.Step(`^saved direct property children remain in schema order$`, func() error { return retainedChildOrder(state) })
	sc.Step(`^the exact changed original member set is (.+) with no additions or removals$`, func(raw string) error {
		var names []string
		if e := json.Unmarshal([]byte(raw), &names); e != nil {
			return e
		}
		if !reflect.DeepEqual(names, state.record.ChangedMembers) {
			return fmt.Errorf("changed member fixture/ledger mismatch")
		}
		return state.custody()
	})
	sc.Step(`^every unpatched selected-property attribute and child retains its literal original bytes$`, func() error { return state.custody() })
	sc.Step(`^all other XML spans, text leaves, runs, paragraphs and unrelated member payloads retain custody$`, func() error { return state.custody() })
	sc.Step(`^actual caller input, original identities, relationships and content-type graph are unchanged$`, func() error { return state.custody() })
	sc.Step(`^refusal reason (.+) preserves session bytes and held target usability before save$`, func(reason string) error {
		if reason != "invalid-style" {
			return fmt.Errorf("unexpected refusal reason")
		}
		var r *packaging.Refusal
		if !errors.As(state.refusal, &r) || r.Kind != reason {
			return fmt.Errorf("refusal=%v want %s", state.refusal, reason)
		}
		if !bytes.Equal(state.caller, state.original) {
			return fmt.Errorf("refusal mutated caller")
		}
		// The same held target must remain usable. Its original line colour is
		// 445566; writing that value must leave the session archive unchanged.
		unchanged := "445566"
		if e := state.pptx.SetRetainedShapeStyle(state.shape, presentation.RetainedShapeStylePatch{LineColor: &unchanged}); e != nil {
			return fmt.Errorf("held style target unusable: %w", e)
		}
		dest := filepath.Join(os.TempDir(), "go-retained-held-"+strings.TrimPrefix(state.record.ID, "@id-")+".pptx")
		defer os.Remove(dest)
		if _, e := state.pptx.SaveAs(dest); e != nil {
			return e
		}
		saved, e := os.ReadFile(dest)
		if e != nil {
			return e
		}
		if !bytes.Equal(saved, state.original) {
			return fmt.Errorf("held refusal changed archive")
		}
		return nil
	})
}
