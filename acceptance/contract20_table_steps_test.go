package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"sort"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

type contract20TableWorld struct {
	source, archive                 []byte
	before, after                   map[string][]byte
	session                         *presentation.EditSession
	target                          *presentation.ContractTableCellTarget
	text, recipe, part, temp, prior string
	result                          error
}

func (w *contract20TableWorld) input(recipe string) error {
	w.recipe = recipe
	w.part = "ppt/slides/slide1.xml"
	w.text = "updated value"
	b, err := os.ReadFile(testutil.ReferencePath("ledgers", "contract20-recipes.json"))
	if err != nil {
		return err
	}
	hash := sha256.Sum256(b)
	if hex.EncodeToString(hash[:]) != contract20Artifacts["ledgers/contract20-recipes.json"] {
		return fmt.Errorf("PPTX recipe seal differs")
	}
	var pack struct {
		PPTX []struct {
			ID, BaseFixtureID string
			Operations        []struct{ Kind, Part, Before, After string }
		}
	}
	if err := json.Unmarshal(b, &pack); err != nil {
		return err
	}
	var recipeRow *struct {
		ID, BaseFixtureID string
		Operations        []struct{ Kind, Part, Before, After string }
	}
	for i := range pack.PPTX {
		if pack.PPTX[i].ID == recipe {
			recipeRow = &pack.PPTX[i]
			break
		}
	}
	if recipeRow == nil {
		return fmt.Errorf("missing PPTX recipe %s", recipe)
	}
	path, err := testutil.LookupFixture(recipeRow.BaseFixtureID)
	if err != nil {
		return err
	}
	base, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	members, err := blankMemberPayloads(base)
	if err != nil {
		return err
	}
	for _, op := range recipeRow.Operations {
		current, ok := members[op.Part]
		if !ok || op.Kind != "replace-literal-once" || strings.Count(string(current), op.Before) != 1 {
			return fmt.Errorf("unsealed PPTX replacement in %s", op.Part)
		}
		members[op.Part] = []byte(strings.Replace(string(current), op.Before, op.After, 1))
	}
	w.source, err = contract20PackMembers(members)
	if err != nil {
		return err
	}
	w.before, err = blankMemberPayloads(w.source)
	if err != nil {
		return err
	}
	w.session, err = presentation.OpenEditing(w.source, packaging.Limits{})
	return err
}
func (w *contract20TableWorld) selectCell() error {
	if w.session == nil {
		return fmt.Errorf("session not open")
	}
	var err error
	w.target, err = w.session.FindContractTableCell(w.part, 4, 0, 0)
	return err
}
func (w *contract20TableWorld) selectedOriginal() error {
	if err := w.selectCell(); err != nil {
		return err
	}
	original := string(w.before[w.part])
	if strings.Count(original, "Galvanic battery") != 1 {
		return fmt.Errorf("original table-cell text not unique")
	}
	var slide struct {
		Frames []struct {
			Pr struct {
				ID string `xml:"id,attr"`
			} `xml:"nvGraphicFramePr>cNvPr"`
			Table struct {
				Rows []struct {
					Cells []struct {
						Body struct {
							Wrap string `xml:"wrap,attr"`
						} `xml:"txBody>bodyPr"`
						Property struct {
							Margin string `xml:"marL,attr"`
							Fill   struct {
								RGB struct {
									Val string `xml:"val,attr"`
								} `xml:"srgbClr"`
							} `xml:"solidFill"`
						} `xml:"tcPr"`
						Paragraph struct {
							PPr struct {
								Align string `xml:"algn,attr"`
							} `xml:"pPr"`
							First struct {
								B    string `xml:"b,attr"`
								Size string `xml:"sz,attr"`
							} `xml:"r>rPr"`
							Text string `xml:"r>t"`
							End  struct {
								Lang string `xml:"lang,attr"`
							} `xml:"endParaRPr"`
						} `xml:"txBody>p"`
					} `xml:"tc"`
				} `xml:"tr"`
			} `xml:"graphic>graphicData>tbl"`
		} `xml:"cSld>spTree>graphicFrame"`
	}
	if err := xml.Unmarshal(w.before[w.part], &slide); err != nil {
		return err
	}
	if len(slide.Frames) != 1 || slide.Frames[0].Pr.ID != "4" || len(slide.Frames[0].Table.Rows) == 0 || len(slide.Frames[0].Table.Rows[0].Cells) == 0 {
		return fmt.Errorf("table frame/cell missing")
	}
	cell := slide.Frames[0].Table.Rows[0].Cells[0]
	if cell.Paragraph.Text != "Galvanic battery" || cell.Body.Wrap != "square" || cell.Property.Margin != "111" || cell.Property.Fill.RGB.Val != "FFFF00" || cell.Paragraph.PPr.Align != "r" || cell.Paragraph.First.B != "1" || cell.Paragraph.First.Size != "1800" || cell.Paragraph.End.Lang != "en-US" {
		return fmt.Errorf("original selected property/text mismatch: %+v", cell)
	}
	return nil
}
func (w *contract20TableWorld) record() error {
	if len(w.before) == 0 || len(w.source) == 0 {
		return fmt.Errorf("no source member snapshot")
	}
	w.prior = filepath.Join(w.temp, "prior.pptx")
	return os.WriteFile(w.prior, []byte("prior-destination"), 0600)
}
func (w *contract20TableWorld) write(text string) error {
	w.text = text
	if w.target == nil {
		return fmt.Errorf("table target not selected")
	}
	w.result = w.session.SetContractTableCellText(w.target, text)
	if w.recipe == "styled-table" {
		return w.result
	}
	return nil
}
func (w *contract20TableWorld) save() error {
	path := filepath.Join(w.temp, "edited.pptx")
	if _, err := w.session.SaveAs(path); err != nil {
		return err
	}
	b, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	w.archive = b
	w.after, err = blankMemberPayloads(b)
	if err != nil {
		return err
	}
	p, err := presentation.Open(path)
	if err != nil {
		return err
	}
	defer p.Close()
	_, err = p.Slide(1)
	return err
}
func (w *contract20TableWorld) custody() error {
	want := "updated value"
	if w.recipe != "styled-table" {
		want = "x"
	}
	if w.text != want {
		return fmt.Errorf("operand changed: %q != %q", w.text, want)
	}
	parts, err := blankMemberPayloads(w.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.before) {
		return fmt.Errorf("caller/source member bytes changed")
	}
	return nil
}
func (w *contract20TableWorld) memberCustody() error {
	if len(w.before) != len(w.after) {
		return fmt.Errorf("member count changed")
	}
	for part, was := range w.before {
		now, ok := w.after[part]
		if !ok {
			return fmt.Errorf("member absent %s", part)
		}
		if part != w.part && !bytes.Equal(was, now) {
			return fmt.Errorf("unselected member %s changed", part)
		}
	}
	return nil
}
func (w *contract20TableWorld) graph() error {
	source, err := packaging.OpenPreserved(w.source, packaging.Limits{})
	if err != nil {
		return err
	}
	saved, err := packaging.OpenPreserved(w.archive, packaging.Limits{})
	if err != nil {
		return err
	}
	before, err := source.Graph()
	if err != nil {
		return err
	}
	after, err := saved.Graph()
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(before, after) {
		return fmt.Errorf("OPC relationship/content-type identities changed")
	}
	for part, old := range w.before {
		if part == "[Content_Types].xml" || strings.HasSuffix(part, ".rels") {
			if !bytes.Equal(old, w.after[part]) {
				return fmt.Errorf("registry changed: %s", part)
			}
		}
	}
	return nil
}
func (w *contract20TableWorld) cell() error {
	original, changed := string(w.before[w.part]), string(w.after[w.part])
	if strings.Count(original, "Galvanic battery") != 1 || changed != strings.Replace(original, "Galvanic battery", w.text, 1) {
		return fmt.Errorf("text-leaf custody failed")
	}
	if strings.Count(changed, "<a:t>updated value</a:t>") != 1 {
		return fmt.Errorf("selected new text absent")
	}
	var xmlCell struct {
		Frames []struct {
			Table struct {
				Rows []struct {
					Cells []struct {
						Body struct {
							Wrap string `xml:"wrap,attr"`
						} `xml:"txBody>bodyPr"`
						Property struct {
							Margin string `xml:"marL,attr"`
							Fill   struct {
								Color string `xml:"val,attr"`
							} `xml:"solidFill>srgbClr"`
						} `xml:"tcPr"`
						Para struct {
							PPr struct {
								Align string `xml:"algn,attr"`
							} `xml:"pPr"`
							Run struct {
								RPr struct {
									B    string `xml:"b,attr"`
									Size string `xml:"sz,attr"`
								} `xml:"rPr"`
								Text string `xml:"t"`
							} `xml:"r"`
							End struct {
								Lang string `xml:"lang,attr"`
							} `xml:"endParaRPr"`
						} `xml:"txBody>p"`
					} `xml:"tc"`
				} `xml:"tr"`
			} `xml:"graphic>graphicData>tbl"`
		} `xml:"cSld>spTree>graphicFrame"`
	}
	if err := xml.Unmarshal(w.after[w.part], &xmlCell); err != nil {
		return err
	}
	if len(xmlCell.Frames) != 1 || len(xmlCell.Frames[0].Table.Rows) == 0 || len(xmlCell.Frames[0].Table.Rows[0].Cells) == 0 {
		return fmt.Errorf("reopened table cell missing")
	}
	cell := xmlCell.Frames[0].Table.Rows[0].Cells[0]
	if cell.Para.Run.Text != w.text || cell.Body.Wrap != "square" || cell.Property.Margin != "111" || cell.Property.Fill.Color != "FFFF00" || cell.Para.PPr.Align != "r" || cell.Para.Run.RPr.B != "1" || cell.Para.Run.RPr.Size != "1800" || cell.Para.End.Lang != "en-US" {
		return fmt.Errorf("saved properties/text drift: %+v", cell)
	}
	return nil
}
func (w *contract20TableWorld) changed() error {
	got := []string{}
	for part, old := range w.before {
		if !bytes.Equal(old, w.after[part]) {
			got = append(got, part)
		}
	}
	sort.Strings(got)
	if !reflect.DeepEqual(got, []string{w.part}) {
		return fmt.Errorf("changed member set %v", got)
	}
	return nil
}
func (w *contract20TableWorld) refusalAndFault() error {
	s, err := presentation.OpenEditing(w.source, packaging.Limits{})
	if err != nil {
		return err
	}
	fresh, err := s.FindContractTableCell(w.part, 4, 0, 0)
	if err != nil {
		return err
	}
	if err = s.SetContractTableCellText(fresh, "\x00"); err == nil {
		return fmt.Errorf("invalid XML text was accepted")
	}
	var refused *packaging.Refusal
	if !errors.As(err, &refused) || refused.Kind != "PPTX_ARGUMENT_INVALID" {
		return fmt.Errorf("invalid XML text refusal %v, want PPTX_ARGUMENT_INVALID", err)
	}
	if err = s.SetContractTableCellText(fresh, "good after refusal"); err != nil {
		return fmt.Errorf("refusal consumed held table cell: %w", err)
	}
	snapshot := func(name string) (map[string][]byte, packaging.Graph, error) {
		path := filepath.Join(w.temp, name)
		if _, e := s.SaveAs(path); e != nil {
			return nil, packaging.Graph{}, e
		}
		archive, e := os.ReadFile(path)
		if e != nil {
			return nil, packaging.Graph{}, e
		}
		members, e := blankMemberPayloads(archive)
		if e != nil {
			return nil, packaging.Graph{}, e
		}
		graph, e := contract20TableGraph(archive)
		return members, graph, e
	}
	baseline, graph, err := snapshot("before-fault.pptx")
	if err != nil {
		return err
	}
	if len(baseline) != len(w.before) || strings.Count(string(baseline[w.part]), "good after refusal") != 1 || string(baseline[w.part]) != strings.Replace(string(w.before[w.part]), "Galvanic battery", "good after refusal", 1) {
		return fmt.Errorf("pre-fault edited session custody differs")
	}
	if err = contract20SaveFaults(filepath.Join(w.temp, "actual-pptx-faults"), func(path string) error { _, e := s.SaveAs(path); return e }); err != nil {
		return err
	}
	after, afterGraph, err := snapshot("after-fault.pptx")
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(baseline, after) || !reflect.DeepEqual(graph, afterGraph) {
		return fmt.Errorf("save fault changed complete edited session members or graph")
	}
	issued, err := s.FindContractTableCell(w.part, 4, 0, 0)
	if err != nil {
		return fmt.Errorf("fault left held cell unselectable: %w", err)
	}
	if err = s.SetContractTableCellText(issued, "good after fault"); err != nil {
		return fmt.Errorf("fault consumed fresh held cell: %w", err)
	}
	final, finalGraph, err := snapshot("after-fault-valid-edit.pptx")
	if err != nil {
		return err
	}
	for part, was := range after {
		if part != w.part && !bytes.Equal(was, final[part]) {
			return fmt.Errorf("valid post-fault edit changed %s", part)
		}
	}
	if !reflect.DeepEqual(graph, finalGraph) || string(final[w.part]) != strings.Replace(string(after[w.part]), "good after refusal", "good after fault", 1) {
		return fmt.Errorf("post-fault valid edit escaped selected cell or changed graph")
	}
	prior, err := os.ReadFile(w.prior)
	if err != nil || string(prior) != "prior-destination" {
		return fmt.Errorf("prior destination overwritten: %v", err)
	}
	return w.custody()
}
func (w *contract20TableWorld) typed() error {
	var refused *packaging.Refusal
	want := "PPTX_TABLE_MERGE_UNSUPPORTED"
	if w.recipe == "malformed-table" {
		want = "PPTX_TABLE_STRUCTURE_UNSUPPORTED"
	}
	if !errors.As(w.result, &refused) || refused.Kind != want {
		return fmt.Errorf("edit refusal %v want %s", w.result, want)
	}
	return nil
}
func (w *contract20TableWorld) refused() error {
	if err := w.custody(); err != nil {
		return err
	}
	if err := w.memberCustody(); err != nil {
		return err
	}
	if err := w.graph(); err != nil {
		return err
	}
	beforeGraph, err := contract20TableGraph(w.archive)
	if err != nil {
		return err
	}
	if err = contract20SaveFaults(filepath.Join(w.temp, "refused-session-faults"), func(path string) error { _, e := w.session.SaveAs(path); return e }); err != nil {
		return err
	}
	path := filepath.Join(w.temp, "refused-after-fault.pptx")
	if _, err = w.session.SaveAs(path); err != nil {
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
	graph, err := contract20TableGraph(archive)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.after) || !reflect.DeepEqual(graph, beforeGraph) || !reflect.DeepEqual(parts, w.before) {
		return fmt.Errorf("refused table session changed after failed saves")
	}
	if err = w.session.SetContractTableCellText(w.target, "x"); err == nil {
		return fmt.Errorf("refused held table cell accepted edit after save fault")
	}
	var refusal *packaging.Refusal
	want := "PPTX_TABLE_MERGE_UNSUPPORTED"
	if w.recipe == "malformed-table" {
		want = "PPTX_TABLE_STRUCTURE_UNSUPPORTED"
	}
	if !errors.As(err, &refusal) || refusal.Kind != want {
		return fmt.Errorf("refused held target after fault %v, want %s", err, want)
	}
	if _, err = w.session.FindContractTableCell(w.part, 4, 0, 0); err != nil {
		return fmt.Errorf("refused session lost selectable target: %w", err)
	}
	prior, err := os.ReadFile(w.prior)
	if err != nil || string(prior) != "prior-destination" {
		return fmt.Errorf("refused session changed prior destination: %v", err)
	}
	return w.custody()
}
func (w *contract20TableWorld) refusedSave() error { return w.save() }
func (w *contract20TableWorld) control() error {
	source := &contract20TableWorld{temp: filepath.Join(w.temp, "control")}
	if err := os.Mkdir(source.temp, 0700); err != nil {
		return err
	}
	if err := source.input("styled-table"); err != nil {
		return err
	}
	if err := source.selectCell(); err != nil {
		return err
	}
	if err := source.write("x"); err != nil {
		return err
	}
	if err := source.save(); err != nil {
		return err
	}
	if err := source.memberCustody(); err != nil {
		return err
	}
	return nil
}

func TestContract20TableSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct{ id, recipe string }{
		{"@id-pptx-table-formatting", "styled-table"},
		{"@id-pptx-table-atomic-refusals", "merged-table"},
	} {
		var selected map[caseID]expectedCase = map[caseID]expectedCase{}
		var key caseID
		for k, v := range inventory {
			if k.ID == tc.id {
				key = k
				selected[k] = v
			}
		}
		count := 1
		if tc.recipe != "styled-table" {
			count = 2
		}
		if len(selected) != count {
			t.Fatalf("selected table case %s has %d rows", tc.id, len(selected))
		}
		t.Run(strings.TrimPrefix(tc.id, "@id-"), func(t *testing.T) {
			state := &contract20TableWorld{temp: t.TempDir()}
			exact := func(sc *godog.ScenarioContext, text string, action func() error) {
				sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
			}
			init := func(sc *godog.ScenarioContext) {
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
					*state = contract20TableWorld{temp: t.TempDir()}
					return ctx, nil
				})
				if tc.recipe == "styled-table" {
					exact(sc, `the contract20 derived PPTX recipe "styled-table" is created in memory from its sealed fixture`, func() error { return state.input("styled-table") })
					exact(sc, `the unique table frame ID 4 cell 0,0 reads JSON "Galvanic battery" with its sealed literal property snapshot`, state.selectedOriginal)
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, state.record)
					exact(sc, `the production retained table-cell editor replaces that cell text with JSON "updated value"`, func() error { return state.write("updated value") })
					exact(sc, `the result is saved to a distinct new path and independently parsed and reopened`, state.save)
					exact(sc, `the source fixture, caller archive and operation operands remain unchanged`, state.custody)
					exact(sc, `only the sealed original member and lexical span allowances differ; all unrelated member payloads remain literal`, state.memberCustody)
					exact(sc, `all saved OPC relationships and content types resolve with original identities preserved`, state.graph)
					exact(sc, `refusals and save faults publish no partial destination and leave the session and held unaffected targets usable`, state.refusalAndFault)
					exact(sc, `the reopened cell text equals JSON "updated value"`, state.cell)
					exact(sc, `the saved properties equal JSON {"bodyPr":{"wrap":"square"},"tcPr":{"marL":"111"},"fill":"FFFF00","paragraph":{"algn":"r"},"firstRun":{"b":"1","sz":"1800"},"endParaRPr":{"lang":"en-US"}}`, state.cell)
					exact(sc, `the exact changed member set is JSON ["ppt/slides/slide1.xml"] without additions or removals`, state.changed)
					exact(sc, `only selected original text leaf content may differ; all bodyPr tcPr pPr rPr endParaRPr attributes children and sibling spans retain literal bytes`, state.cell)
				} else {
					for _, r := range []string{"merged-table", "malformed-table"} {
						recipe := r
						exact(sc, fmt.Sprintf(`the contract20 derived PPTX recipe "%s" is created in memory from its sealed fixture`, r), func() error { return state.input(recipe) })
					}
					exact(sc, `the source and prior destination bytes and all member payloads are recorded`, state.record)
					exact(sc, `the production package open and target selection of table frame ID 4 cell 0,0 succeed without mutation or intake refusal`, state.selectCell)
					exact(sc, `the production table-cell editor attempts JSON "x" at the selected target`, func() error { return state.write("x") })
					for _, row := range []struct{ code string }{{"PPTX_TABLE_MERGE_UNSUPPORTED"}, {"PPTX_TABLE_STRUCTURE_UNSUPPORTED"}} {
						code := row.code
						exact(sc, fmt.Sprintf(`only that cell-text edit preflight returns exactly typed refusal category "%s" and no changed result`, code), state.typed)
					}
					exact(sc, `every source member, selected table property and relationship remains unchanged`, state.custody)
					exact(sc, `saving the refused session retains all original members without additions or removals`, state.refusedSave)
					exact(sc, `no refusal or save fault changes source caller operands or any prior destination`, state.refused)
					exact(sc, `an independently constructed unmerged same-shape control permits the identical cell mutation and retains unrelated bytes`, state.control)
				}
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-table-" + strings.TrimPrefix(tc.id, "@"), ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: tc.id, Strict: true, Concurrency: 1}}
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
			if err := reconcile(selected, data); err != nil {
				t.Error(err)
			}
			if dir := os.Getenv("GO_CONTRACT20_PPTX_TABLE_RECEIPTS_DIR"); dir != "" && !t.Failed() {
				if e := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(tc.id, "@id-")+".cucumber.json"), output.Bytes(), 0600); e != nil {
					t.Fatal(e)
				}
			}
		})
	}
}
