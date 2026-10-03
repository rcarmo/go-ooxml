package acceptance

import (
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

type contract20FormulaRefusalWorld struct {
	input         contract20ReadWorld
	session       *spreadsheet.EditSession
	held          *spreadsheet.NumberTarget
	address, kind string
	destination   string
	operand       float64
	result        error
}

func (w *contract20FormulaRefusalWorld) recipe() error {
	if err := w.input.input("attributed-formulas"); err != nil {
		return err
	}
	// Parse every anchor, follower and cache in the literal shared recipe;
	// substring presence alone could falsely accept a misplaced formula.
	var sheet struct {
		Rows []struct {
			Cells []struct {
				Address string `xml:"r,attr"`
				Formula *struct {
					Kind  string `xml:"t,attr"`
					Index string `xml:"si,attr"`
					Range string `xml:"ref,attr"`
					Text  string `xml:",chardata"`
				} `xml:"f"`
				Value *string `xml:"v"`
			} `xml:"c"`
		} `xml:"sheetData>row"`
	}
	if err := xml.Unmarshal(w.input.parts["xl/worksheets/sheet1.xml"], &sheet); err != nil {
		return err
	}
	cells := map[string]struct {
		formula *struct{ Kind, Index, Range, Text string }
		cache   *string
	}{}
	for _, row := range sheet.Rows {
		for _, cell := range row.Cells {
			if _, exists := cells[cell.Address]; exists {
				return fmt.Errorf("duplicate formula fixture cell %s", cell.Address)
			}
			var formula *struct{ Kind, Index, Range, Text string }
			if cell.Formula != nil {
				formula = &struct{ Kind, Index, Range, Text string }{cell.Formula.Kind, cell.Formula.Index, cell.Formula.Range, cell.Formula.Text}
			}
			cells[cell.Address] = struct {
				formula *struct{ Kind, Index, Range, Text string }
				cache   *string
			}{formula, cell.Value}
		}
	}
	for _, want := range []struct{ address, kind, index, span, body, cache string }{
		{"B2", "shared", "0", "B2:B3", "A2*2", "2"}, {"B3", "shared", "0", "", "", "4"}, {"D2", "array", "", "D2:D3", "A2:A3*2", "2"}, {"D3", "", "", "", "", "4"},
	} {
		c, ok := cells[want.address]
		if !ok || c.cache == nil || *c.cache != want.cache {
			return fmt.Errorf("missing formula/follower/cache %s", want.address)
		}
		if want.kind == "" {
			if c.formula != nil {
				return fmt.Errorf("unexpected follower formula %s", want.address)
			}
		} else if c.formula == nil || c.formula.Kind != want.kind || c.formula.Index != want.index || c.formula.Range != want.span || c.formula.Text != want.body {
			return fmt.Errorf("attributed formula mismatch %s: %+v", want.address, c.formula)
		}
	}
	return nil
}
func (w *contract20FormulaRefusalWorld) record(temp string) error {
	if len(w.input.parts) == 0 || !bytes.Equal(w.input.original, w.input.source) {
		return fmt.Errorf("source not recorded")
	}
	w.destination = filepath.Join(temp, "refused.xlsx")
	var err error
	w.session, err = spreadsheet.OpenEditing(w.input.source, packaging.Limits{})
	if err != nil {
		return err
	}
	w.held, err = w.session.FindNumber("Calc", "A2")
	return err
}
func (w *contract20FormulaRefusalWorld) attempt(address, kind string) error {
	w.address, w.kind, w.operand = address, kind, 99
	// Refusal is the outcome of the production editor, not FindNumber's
	// classification or an acceptance-local gate.
	w.result = w.session.SetNumberAt("Calc", address, w.operand)
	return nil
}
func (w *contract20FormulaRefusalWorld) typed() error {
	var refused *packaging.Refusal
	if !errors.As(w.result, &refused) || refused.Kind != w.kind {
		return fmt.Errorf("write %s = %v, expected typed %s", w.address, w.result, w.kind)
	}
	return nil
}
func (w *contract20FormulaRefusalWorld) unchanged() error {
	if !bytes.Equal(w.input.original, w.input.source) {
		return fmt.Errorf("caller archive mutated")
	}
	current, err := blankMemberPayloads(w.input.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(current, w.input.parts) {
		return fmt.Errorf("original formula/follower/cache member changed")
	}
	return nil
}
func (w *contract20FormulaRefusalWorld) saveRefused() error {
	if _, err := os.Stat(w.destination); !os.IsNotExist(err) {
		return fmt.Errorf("unexpected prior refused result: %v", err)
	}
	if _, err := w.session.SaveAs(w.destination); err != nil {
		return err
	}
	output, err := os.ReadFile(w.destination)
	if err != nil {
		return err
	}
	parts, err := blankMemberPayloads(output)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.input.parts) {
		return fmt.Errorf("refused save changed member payloads or names")
	}
	original, err := packaging.OpenPreserved(w.input.original, packaging.Limits{})
	if err != nil {
		return err
	}
	saved, err := packaging.OpenPreserved(output, packaging.Limits{})
	if err != nil {
		return err
	}
	before, err := original.Graph()
	if err != nil {
		return err
	}
	after, err := saved.Graph()
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(before, after) {
		return fmt.Errorf("refused save changed package graph")
	}
	return nil
}
func (w *contract20FormulaRefusalWorld) deliveryCustody() error {
	if w.operand != 99 || w.address != "B2" && w.address != "D2" {
		return fmt.Errorf("caller operands changed")
	}
	if err := w.unchanged(); err != nil {
		return err
	}
	if _, err := os.Stat(w.destination); err != nil {
		return fmt.Errorf("new refused-save destination missing: %w", err)
	}
	if err := contract20SaveFaults(filepath.Dir(w.destination), func(path string) error {
		_, saveErr := w.session.SaveAs(path)
		return saveErr
	}); err != nil {
		return err
	}
	if err := w.heldValue(); err != nil {
		return err
	}
	// A fresh post-fault save proves the retained session still carries all
	// original member payloads, not merely the caller archive snapshot.
	post := filepath.Join(filepath.Dir(w.destination), "after-fault.xlsx")
	if _, err := w.session.SaveAs(post); err != nil {
		return err
	}
	data, err := os.ReadFile(post)
	if err != nil {
		return err
	}
	parts, err := blankMemberPayloads(data)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.input.parts) {
		return fmt.Errorf("retained refused session payloads changed after faults")
	}
	return w.unchanged()
}
func (w *contract20FormulaRefusalWorld) heldValue() error {
	value, err := w.session.CurrentNumber(w.held)
	if err != nil || value != 1 {
		return fmt.Errorf("held Calc!A2 after refusal: %v/%v", value, err)
	}
	return nil
}
func (w *contract20FormulaRefusalWorld) heldUsable() error {
	if err := w.heldValue(); err != nil {
		return err
	}
	// The held target remains live for an independent no-op observation.
	effect, err := w.session.SetNumberWithAllCachesInvalidated(w.held, 1)
	if err != nil || effect.ValueChanged || len(effect.Invalidated) != 0 {
		return fmt.Errorf("held no-op = %+v/%v", effect, err)
	}
	return w.unchanged()
}

func TestContract20FormulaRefusalSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct{ id, address, kind string }{
		{"@id-xlsx-refuse-shared-formula-overwrite", "B2", "xlsx-shared-formula-edit-unsupported"},
		{"@id-xlsx-refuse-array-formula-overwrite", "D2", "xlsx-array-formula-edit-unsupported"},
	} {
		var key caseID
		var want expectedCase
		count := 0
		for k, v := range inventory {
			if k.ID == tc.id {
				key, want = k, v
				count++
			}
		}
		if count != 1 {
			t.Fatalf("selected formula case %s has %d rows", tc.id, count)
		}
		t.Run(strings.TrimPrefix(tc.id, "@id-"), func(t *testing.T) {
			state := &contract20FormulaRefusalWorld{}
			temp := t.TempDir()
			exact := func(sc *godog.ScenarioContext, text string, action func() error) {
				sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
			}
			init := func(sc *godog.ScenarioContext) {
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
					*state = contract20FormulaRefusalWorld{}
					return ctx, nil
				})
				exact(sc, `the literal contract20 XLSX recipe "attributed-formulas" is packed without workbook APIs`, state.recipe)
				exact(sc, `the source and prior destination bytes, every member and held input-cell target are recorded`, func() error { return state.record(temp) })
				exact(sc, fmt.Sprintf(`the production value editor attempts numeric 99 at Calc!%s`, tc.address), func() error { return state.attempt(tc.address, tc.kind) })
				exact(sc, fmt.Sprintf(`the operation returns exactly typed refusal category "%s" and no edited result`, tc.kind), state.typed)
				exact(sc, `the original formula anchors followers caches and all member payloads remain unchanged`, state.unchanged)
				exact(sc, `saving the refused session to a new path preserves every original member without additions or removals`, state.saveRefused)
				exact(sc, `source caller operands and prior destination remain unchanged with no partial output`, state.deliveryCustody)
				exact(sc, `a held unrelated Calc!A2 target remains usable after the refusal`, state.heldUsable)
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-formula-" + strings.TrimPrefix(tc.id, "@"), ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: tc.id, Strict: true, Concurrency: 1}}
			if code := suite.Run(); code != 0 {
				t.Errorf("Godog %d: %s", code, output.String())
			}
			var report []reportFeature
			if err := json.Unmarshal(output.Bytes(), &report); err != nil {
				t.Fatalf("report: %v: %s", err, output.String())
			}
			data, err := json.Marshal(report)
			if err != nil {
				t.Fatal(err)
			}
			if err := reconcile(map[caseID]expectedCase{key: want}, data); err != nil {
				t.Error(err)
			}
			if dir := os.Getenv("GO_CONTRACT20_XLSX_RECEIPTS_DIR"); dir != "" && !t.Failed() {
				if err := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(tc.id, "@id-")+".cucumber.json"), output.Bytes(), 0600); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
}
