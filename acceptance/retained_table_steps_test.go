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

	godog "github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type tableState struct {
	record                  tableRecord
	caller, original, saved []byte
	before, after           map[string][]byte
	pptx                    *presentation.EditSession
	word                    *document.EditSession
	frame                   *presentation.RetainedTableTarget
	cell                    *document.RetainedTableTarget
	refusal                 error
}

func (s *tableState) part() string {
	if s.record.Format == "pptx" {
		return "ppt/slides/slide1.xml"
	}
	return "word/document.xml"
}
func (s *tableState) perform() error {
	var patch map[string]any
	if e := json.Unmarshal(s.record.Patch, &patch); e != nil {
		return e
	}
	if patch == nil {
		return fmt.Errorf("empty table patch")
	}
	if s.pptx != nil {
		return s.pptx.SetRetainedTable(s.frame, patch)
	}
	return s.word.SetRetainedTable(s.cell, patch)
}
func (s *tableState) custody() error {
	if s.saved == nil || s.after == nil {
		return fmt.Errorf("saved output absent")
	}
	if !bytes.Equal(s.caller, s.original) {
		return fmt.Errorf("caller buffer drift")
	}
	if len(s.before) != len(s.after) {
		return fmt.Errorf("member count drift")
	}
	actual := map[string]bool{}
	for name, b := range s.before {
		now, ok := s.after[name]
		if !ok {
			return fmt.Errorf("member removed %s", name)
		}
		if !bytes.Equal(b, now) {
			actual[name] = true
		}
	}
	expected := map[string]bool{}
	for _, name := range s.record.ChangedMembers {
		expected[name] = true
	}
	if !reflect.DeepEqual(actual, expected) {
		return fmt.Errorf("changed members %v want %v", actual, expected)
	}
	a, e := packaging.OpenPreserved(s.original, packaging.Limits{})
	if e != nil {
		return e
	}
	b, e := packaging.OpenPreserved(s.saved, packaging.Limits{})
	if e != nil {
		return e
	}
	ag, e := a.Graph()
	if e != nil {
		return e
	}
	bg, e := b.Graph()
	if e != nil {
		return e
	}
	if !reflect.DeepEqual(ag, bg) {
		return fmt.Errorf("OPC graph drift")
	}
	if s.record.Kind == "refusal" && !bytes.Equal(s.saved, s.original) {
		return fmt.Errorf("refusal archive drift")
	}
	if len(expected) > 0 {
		return tableMaskedSame(s.before[s.part()], s.after[s.part()], s.record)
	}
	return nil
}
func tableSteps(sc *godog.ScenarioContext) {
	state := &tableState{}
	sc.Before(func(ctx context.Context, scenario *godog.Scenario) (context.Context, error) {
		*state = tableState{}
		records, e := tableRecords()
		if e != nil {
			return ctx, e
		}
		for _, tag := range scenario.Tags {
			if r, ok := records[tag.Name]; ok {
				state.record = r
				break
			}
		}
		// This initializer also runs for the historical native suite.
		return ctx, nil
	})
	sc.Step(`^retained (pptx|docx) input (fixture-[0-9a-f]+) has its sealed original property and identity snapshot$`, func(format, id string) error {
		if format != state.record.Format || id != state.record.FixtureID {
			return fmt.Errorf("table fixture record mismatch")
		}
		path, e := testutil.LookupFixture(id)
		if e != nil {
			return e
		}
		b, e := os.ReadFile(path)
		if e != nil {
			return e
		}
		sum := sha256.Sum256(b)
		if "fixture-"+hex.EncodeToString(sum[:]) != id {
			return fmt.Errorf("table fixture provenance")
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
			return fmt.Errorf("target mismatch %s", target)
		}
		var e error
		if state.pptx != nil {
			state.frame, e = state.pptx.FindRetainedTable(state.part(), 3, 0, 0)
		} else {
			state.cell, e = state.word.FindRetainedTable(0, 0, 0)
		}
		return e
	})
	sc.Step(`^production retained (pptx|docx) editing applies ([a-z-]+) patch JSON (.+)$`, func(format, kind, raw string) error {
		if format != state.record.Format || kind != state.record.Kind {
			return fmt.Errorf("table operation mismatch")
		}
		if e := retainedJSON(raw, state.record.Patch); e != nil {
			return e
		}
		e := state.perform()
		if state.record.Kind == "refusal" {
			state.refusal = e
			return nil
		}
		return e
	})
	sc.Step(`^the result or unchanged refusal session is saved and independently parsed and reopened$`, func() error {
		dest := filepath.Join(os.TempDir(), "go-table-"+strings.TrimPrefix(state.record.ID, "@id-")+"-output."+state.record.Format)
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
			return fmt.Errorf("saved target mismatch")
		}
		if e := retainedJSON(raw, state.record.Expected); e != nil {
			return e
		}
		return tableExpected(state)
	})
	sc.Step(`^saved direct property children remain in schema order$`, func() error { return tableOrder(state) })
	sc.Step(`^the exact changed original member set is (.+) with no additions or removals$`, func(raw string) error {
		var names []string
		if e := json.Unmarshal([]byte(raw), &names); e != nil {
			return e
		}
		if !reflect.DeepEqual(names, state.record.ChangedMembers) {
			return fmt.Errorf("changed member ledger mismatch")
		}
		return state.custody()
	})
	sc.Step(`^every unpatched selected-property attribute and child retains its literal original bytes$`, func() error { return state.custody() })
	sc.Step(`^all other XML spans, text leaves, runs, paragraphs and unrelated member payloads retain custody$`, func() error { return state.custody() })
	sc.Step(`^actual caller input, original identities, relationships and content-type graph are unchanged$`, func() error { return state.custody() })
	sc.Step(`^refusal reason (.+) preserves session bytes and held target usability before save$`, func(reason string) error {
		var r *packaging.Refusal
		if reason != "invalid-table-properties" || !errors.As(state.refusal, &r) || r.Kind != reason {
			return fmt.Errorf("table refusal: %v", state.refusal)
		}
		if e := state.custody(); e != nil {
			return e
		}
		if state.pptx != nil {
			return state.pptx.SetRetainedTable(state.frame, map[string]any{"fill": "4472C4"})
		}
		return state.word.SetRetainedTable(state.cell, map[string]any{"shading": "4472C4"})
	})
}

var _ = losslessxml.Parse
