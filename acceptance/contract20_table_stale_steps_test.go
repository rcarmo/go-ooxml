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
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type contract20StaleTableWorld struct {
	source  contract20TableWorld
	old     *presentation.ContractTableCellTarget
	frameID uint32
	added   map[string][]byte
	graph   packaging.Graph
	prior   string
	attempt error
}

const contract20StaleTablePart = "ppt/slides/slide1.xml"

func (w *contract20StaleTableWorld) input() error {
	if err := w.source.input("stale-table"); err != nil {
		return err
	}
	var e error
	w.old, e = w.source.session.FindContractTableCell(contract20StaleTablePart, 4, 0, 0)
	if e != nil {
		return e
	}
	old := string(w.source.before[contract20StaleTablePart])
	if strings.Count(old, `<p:cNvPr id="4"`) != 1 || strings.Count(old, `<a:p></a:p>`) != 1 {
		return fmt.Errorf("original frame/blank cell not uniquely sealed")
	}
	return nil
}
func (w *contract20StaleTableWorld) add() error {
	var err error
	w.frameID, err = w.source.session.AddContractTable(contract20StaleTablePart, 1, 1, 1200, 100, 900, 600)
	if err != nil {
		return err
	}
	if w.frameID == 0 || w.frameID == 4 {
		return fmt.Errorf("allocated frame identity %d", w.frameID)
	}
	return nil
}
func (w *contract20StaleTableWorld) record() error {
	temp := w.source.temp
	w.prior = filepath.Join(temp, "prior-destination.pptx")
	if err := os.WriteFile(w.prior, []byte("untouched prior destination"), 0600); err != nil {
		return err
	}
	path := filepath.Join(temp, "post-add.pptx")
	if _, err := w.source.session.SaveAs(path); err != nil {
		return err
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	w.added, err = blankMemberPayloads(data)
	if err != nil {
		return err
	}
	w.graph, err = contract20TableGraph(data)
	if err != nil {
		return err
	}
	originalGraph, err := contract20TableGraph(w.source.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(w.graph, originalGraph) {
		return fmt.Errorf("table insertion changed OPC graph identities")
	}
	if len(w.added) != len(w.source.before) {
		return fmt.Errorf("post-add member inventory drift")
	}
	for part, old := range w.source.before {
		now, ok := w.added[part]
		if !ok || part != contract20StaleTablePart && !bytes.Equal(old, now) {
			return fmt.Errorf("post-add unselected part %s changed", part)
		}
	}
	// Erase only the uniquely identified new direct frame; the source and every
	// untouched lexical byte must then be identical, not just semantically equal.
	doc, err := losslessxml.Parse(w.added[contract20StaleTablePart])
	if err != nil {
		return err
	}
	es := doc.Elements()
	var inserted losslessxml.Element
	count := 0
	for _, node := range es {
		if node.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "graphicFrame"}) {
			continue
		}
		for _, n := range es {
			p, ok := n.Parent()
			if ok && p.Name() == (xml.Name{Space: packaging.NSPresentationML, Local: "nvGraphicFramePr"}) && n.Name() == (xml.Name{Space: packaging.NSPresentationML, Local: "cNvPr"}) {
				grand, ok := p.Parent()
				if !ok || grand != node {
					continue
				}
				for _, a := range n.Attributes() {
					if a.Name.Local == "id" && a.Value == fmt.Sprint(w.frameID) {
						inserted = node
						count++
					}
				}
			}
		}
	}
	if count != 1 {
		return fmt.Errorf("inserted frame identity count %d", count)
	}
	a, b := inserted.SourceRange()
	post := w.added[contract20StaleTablePart]
	without := append(bytes.Clone(post[:a]), post[b:]...)
	if !bytes.Equal(without, w.source.before[contract20StaleTablePart]) {
		return fmt.Errorf("post-add modified original slide spans")
	}
	return nil
}
func (w *contract20StaleTableWorld) writeOld() error {
	w.attempt = w.source.session.SetContractTableCellText(w.old, "stale write")
	return nil
}
func (w *contract20StaleTableWorld) typed() error {
	var refusal *packaging.Refusal
	if !errors.As(w.attempt, &refusal) || refusal.Kind != "PPTX_STALE_TABLE_HANDLE" {
		return fmt.Errorf("old handle write = %v, want PPTX_STALE_TABLE_HANDLE", w.attempt)
	}
	return nil
}
func (w *contract20StaleTableWorld) custody() error {
	source, err := blankMemberPayloads(w.source.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(source, w.source.before) {
		return fmt.Errorf("caller source changed")
	}
	path := filepath.Join(w.source.temp, "after-stale.pptx")
	if _, err := w.source.session.SaveAs(path); err != nil {
		return err
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	parts, err := blankMemberPayloads(data)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.added) {
		return fmt.Errorf("stale write changed post-add member state")
	}
	graph, err := contract20TableGraph(data)
	if err != nil || !reflect.DeepEqual(graph, w.graph) {
		return fmt.Errorf("stale write changed OPC graph: %v", err)
	}
	prior, err := os.ReadFile(w.prior)
	if err != nil || string(prior) != "untouched prior destination" {
		return fmt.Errorf("prior destination changed %q/%v", prior, err)
	}
	return nil
}
func (w *contract20StaleTableWorld) reopen() error {
	path := filepath.Join(w.source.temp, "reopened.pptx")
	if _, err := w.source.session.SaveAs(path); err != nil {
		return err
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	opened, err := presentation.Open(path)
	if err != nil {
		return err
	}
	defer opened.Close()
	first, err := opened.Slide(1)
	if err != nil {
		return err
	}
	if len(first.Tables()) != 2 {
		return fmt.Errorf("reopened table count %d", len(first.Tables()))
	}
	for i, table := range first.Tables() {
		if table.RowCount() != 1 || table.ColumnCount() != 1 || table.Cell(0, 0).Text() != "" {
			return fmt.Errorf("reopened table %d not blank one-by-one", i)
		}
	}
	parts, err := blankMemberPayloads(data)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.added) {
		return fmt.Errorf("reopened state changed after stale write")
	}
	return nil
}
func (w *contract20StaleTableWorld) fresh() error {
	fresh, err := w.source.session.FindContractTableCell(contract20StaleTablePart, 4, 0, 0)
	if err != nil {
		return err
	}
	// An invalid late write must refuse before touching this freshly issued
	// target, and re-selection of both frame IDs must still succeed.
	err = w.source.session.SetContractTableCellText(fresh, "\x00")
	var refusal *packaging.Refusal
	if !errors.As(err, &refusal) || refusal.Kind != "PPTX_ARGUMENT_INVALID" {
		return fmt.Errorf("invalid late write %v", err)
	}
	if _, err = w.source.session.FindContractTableCell(contract20StaleTablePart, 4, 0, 0); err != nil {
		return err
	}
	if _, err = w.source.session.FindContractTableCell(contract20StaleTablePart, w.frameID, 0, 0); err != nil {
		return err
	}
	if err := w.custody(); err != nil {
		return err
	}
	// An issued target remains usable after refusal: now commit one valid edit
	// and independently verify only the original first-table text subtree moved.
	if err := w.source.session.SetContractTableCellText(fresh, "usable"); err != nil {
		return fmt.Errorf("fresh handle not usable after refusal: %w", err)
	}
	path := filepath.Join(w.source.temp, "fresh-edited.pptx")
	if _, err := w.source.session.SaveAs(path); err != nil {
		return err
	}
	archive, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	parts, err := blankMemberPayloads(archive)
	if err != nil {
		return err
	}
	for part, before := range w.added {
		if part != contract20StaleTablePart && !bytes.Equal(before, parts[part]) {
			return fmt.Errorf("fresh write changed unrelated part %s", part)
		}
	}
	if strings.Count(string(w.added[contract20StaleTablePart]), `<a:p></a:p>`) != 1 {
		return fmt.Errorf("original empty cell not uniquely identified")
	}
	want := bytes.Replace(w.added[contract20StaleTablePart], []byte(`<a:p></a:p>`), []byte(`<a:p><a:r><a:t>usable</a:t></a:r></a:p>`), 1)
	if !bytes.Equal(parts[contract20StaleTablePart], want) {
		return fmt.Errorf("fresh write escaped original cell lexical boundary")
	}
	return nil
}

func contract20TableGraph(data []byte) (packaging.Graph, error) {
	p, err := packaging.OpenPreserved(data, packaging.Limits{})
	if err != nil {
		return packaging.Graph{}, err
	}
	return p.Graph()
}

func TestContract20StaleTableSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	const id = "@id-pptx-table-stale-handle"
	var key caseID
	var want expectedCase
	count := 0
	for k, v := range inventory {
		if k.ID == id {
			key, want = k, v
			count++
		}
	}
	if count != 1 {
		t.Fatalf("selected table case rows %d", count)
	}
	state := &contract20StaleTableWorld{source: contract20TableWorld{temp: t.TempDir()}}
	exact := func(sc *godog.ScenarioContext, text string, action func() error) {
		sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
	}
	init := func(sc *godog.ScenarioContext) {
		sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
			*state = contract20StaleTableWorld{source: contract20TableWorld{temp: t.TempDir()}}
			return ctx, nil
		})
		exact(sc, `the contract20 derived PPTX recipe "stale-table" is created in memory from its sealed fixture`, state.input)
		exact(sc, `the production editor holds table frame ID 4 cell 0,0 with its original empty text`, func() error {
			if state.old == nil {
				return fmt.Errorf("table handle absent")
			}
			return nil
		})
		exact(sc, `the production editor adds a second 1-row 1-column table at EMU x=1200 y=100 width=900 height=600`, state.add)
		exact(sc, `the complete post-add member state and prior destination are recorded`, state.record)
		exact(sc, `the old first-table cell handle attempts JSON "stale write"`, state.writeOld)
		exact(sc, `the operation returns exactly typed refusal category "PPTX_STALE_TABLE_HANDLE" and no changed result`, state.typed)
		exact(sc, `every post-add member payload and graph identity, caller input and prior destination remains unchanged`, state.custody)
		exact(sc, `saving and independently reopening still finds exactly two tables with both first cells empty`, state.reopen)
		exact(sc, `a freshly found first-table cell is usable and any invalid late patch refuses without consuming it`, state.fresh)
	}
	var output bytes.Buffer
	suite := godog.TestSuite{Name: "go-contract20-stale-table", ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: id, Strict: true, Concurrency: 1}}
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
	if dir := os.Getenv("GO_CONTRACT20_PPTX_TABLE_RECEIPTS_DIR"); dir != "" && !t.Failed() {
		if e := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(id, "@id-")+".cucumber.json"), output.Bytes(), 0600); e != nil {
			t.Fatal(e)
		}
	}
}
