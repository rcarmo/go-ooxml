package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"sort"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

type contract20BlankWorld struct {
	read        contract20ReadWorld
	session     *spreadsheet.EditSession
	sheet, part string
	edits       []spreadsheet.ContractBlankEdit
	output      map[string][]byte
	archive     []byte
	path        string
	temp        string
}

func (w *contract20BlankWorld) input(recipe string) error {
	if err := w.read.input(recipe); err != nil {
		return err
	}
	w.part = "xl/worksheets/sheet1.xml"
	if recipe == "prefixed-cells" {
		w.sheet = "Prefixed"
		w.edits = []spreadsheet.ContractBlankEdit{{Address: "A1", Value: "Alpha"}, {Address: "B1", Value: float64(7)}}
	} else {
		w.sheet = "StyledBlank"
		w.edits = []spreadsheet.ContractBlankEdit{{Address: "A1", Value: "filled"}}
	}
	if _, err := w.styles(); err != nil {
		return err
	}
	return nil
}
func (w *contract20BlankWorld) styles() (map[string]string, error) {
	var table struct {
		CellXfs struct {
			Count string     `xml:"count,attr"`
			XFs   []struct{} `xml:"xf"`
		} `xml:"cellXfs"`
	}
	if err := xml.Unmarshal(w.read.parts["xl/styles.xml"], &table); err != nil {
		return nil, err
	}
	if table.CellXfs.Count != "3" || len(table.CellXfs.XFs) != 3 {
		return nil, fmt.Errorf("sealed registry 0/1/2 unavailable")
	}
	return cellAttributes(w.read.parts[w.part])
}
func cellAttributes(data []byte) (map[string]string, error) {
	var worksheet struct {
		Rows []struct {
			Cells []struct {
				R string `xml:"r,attr"`
				S string `xml:"s,attr"`
			} `xml:"c"`
		} `xml:"sheetData>row"`
	}
	if err := xml.Unmarshal(data, &worksheet); err != nil {
		return nil, err
	}
	attrs := map[string]string{}
	for _, row := range worksheet.Rows {
		for _, cell := range row.Cells {
			if cell.R == "" || attrs[cell.R] != "" {
				return nil, fmt.Errorf("missing/duplicate cell reference %q", cell.R)
			}
			attrs[cell.R] = cell.S
		}
	}
	return attrs, nil
}
func (w *contract20BlankWorld) beforeStyled() error {
	wb, err := spreadsheet.OpenReader(bytes.NewReader(w.read.source), int64(len(w.read.source)))
	if err != nil {
		return err
	}
	defer wb.Close()
	sheet, err := wb.Sheet("StyledBlank")
	if err != nil {
		return err
	}
	if sheet.Cell("A1").Type() != spreadsheet.CellTypeEmpty || sheet.Cell("A1").Value() != nil {
		return fmt.Errorf("A1 not null before edit")
	}
	if v, e := sheet.Cell("B1").Float64(); e != nil || v != 5 {
		return fmt.Errorf("B1 prior %v/%v", v, e)
	}
	styles, err := w.styles()
	if err != nil {
		return err
	}
	if styles["A1"] != "1" || styles["B1"] != "2" {
		return fmt.Errorf("source style indexes: %v", styles)
	}
	return nil
}
func (w *contract20BlankWorld) record() error {
	if len(w.read.parts) == 0 || len(w.read.source) == 0 || !bytes.Equal(w.read.source, w.read.original) {
		return fmt.Errorf("source not recorded")
	}
	return nil
}
func (w *contract20BlankWorld) write() error {
	s, err := spreadsheet.OpenEditing(w.read.source, packaging.Limits{})
	if err != nil {
		return err
	}
	w.session = s
	return s.SetContractBlankValues(w.sheet, w.edits)
}
func (w *contract20BlankWorld) save() error {
	if w.session == nil {
		return fmt.Errorf("no edited session")
	}
	w.path = filepath.Join(w.temp, "saved.xlsx")
	if _, err := w.session.SaveAs(w.path); err != nil {
		return err
	}
	b, err := os.ReadFile(w.path)
	if err != nil {
		return err
	}
	w.archive = b
	// Independently inspect the saved ZIP rather than trusting mutation receipts.
	w.output, err = blankMemberPayloads(b)
	if err != nil {
		return err
	}
	opened, err := spreadsheet.OpenReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return err
	}
	defer opened.Close()
	_, err = opened.Sheet(w.sheet)
	return err
}
func blankMemberPayloads(data []byte) (map[string][]byte, error) {
	archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	parts := map[string][]byte{}
	for _, file := range archive.File {
		if _, exists := parts[file.Name]; exists {
			return nil, fmt.Errorf("duplicate member %s", file.Name)
		}
		reader, e := file.Open()
		if e != nil {
			return nil, e
		}
		value, e := io.ReadAll(reader)
		closeErr := reader.Close()
		if e != nil {
			return nil, e
		}
		if closeErr != nil {
			return nil, closeErr
		}
		parts[file.Name] = value
	}
	return parts, nil
}
func (w *contract20BlankWorld) inputCustody() error {
	if !bytes.Equal(w.read.source, w.read.original) {
		return fmt.Errorf("caller archive changed")
	}
	for _, edit := range w.edits {
		if edit.Address == "A1" {
			if w.sheet == "Prefixed" && edit.Value != "Alpha" || w.sheet == "StyledBlank" && edit.Value != "filled" {
				return fmt.Errorf("operand changed")
			}
		} else if edit.Address != "B1" || edit.Value != float64(7) {
			return fmt.Errorf("operand changed")
		}
	}
	return w.read.custody()
}
func (w *contract20BlankWorld) memberCustody() error {
	if len(w.output) != len(w.read.parts) {
		return fmt.Errorf("member count %d -> %d", len(w.read.parts), len(w.output))
	}
	for part, original := range w.read.parts {
		changed := part == w.part || w.sheet == "Prefixed" && part == "xl/workbook.xml"
		now, ok := w.output[part]
		if !ok {
			return fmt.Errorf("missing member %s", part)
		}
		if !changed && !bytes.Equal(original, now) {
			return fmt.Errorf("unselected member changed %s", part)
		}
	}
	return nil
}
func (w *contract20BlankWorld) relationships() error {
	original, err := packaging.OpenPreserved(w.read.original, packaging.Limits{})
	if err != nil {
		return err
	}
	current, err := packaging.OpenPreserved(w.archive, packaging.Limits{})
	if err != nil {
		return err
	}
	prior, err := original.Graph()
	if err != nil {
		return err
	}
	after, err := current.Graph()
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(prior, after) {
		return fmt.Errorf("OPC parts/relationships/content types differ")
	}
	for _, name := range []string{"[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels"} {
		if !bytes.Equal(w.read.parts[name], w.output[name]) {
			return fmt.Errorf("registry changed: %s", name)
		}
	}
	return nil
}
func (w *contract20BlankWorld) values() error {
	opened, err := spreadsheet.OpenReader(bytes.NewReader(w.archive), int64(len(w.archive)))
	if err != nil {
		return err
	}
	defer opened.Close()
	sheet, err := opened.Sheet(w.sheet)
	if err != nil {
		return err
	}
	want := "filled"
	if w.sheet == "Prefixed" {
		want = "Alpha"
	}
	if got := sheet.Cell("A1").String(); got != want {
		return fmt.Errorf("reopened A1 %q != %q", got, want)
	}
	styles, err := cellAttributes(w.output[w.part])
	if err != nil {
		return err
	}
	if styles["A1"] != "1" || styles["B1"] != "2" {
		return fmt.Errorf("reopened styles %v", styles)
	}
	number := float64(5)
	if w.sheet == "Prefixed" {
		number = 7
	}
	if got, e := sheet.Cell("B1").Float64(); e != nil || got != number {
		return fmt.Errorf("reopened B1 %v/%v", got, e)
	}
	return nil
}
func (w *contract20BlankWorld) qualified() error {
	text := string(w.output[w.part])
	for _, fragment := range []string{`<x:is><x:t>Alpha</x:t></x:is>`, `<x:v>7</x:v>`} {
		if strings.Count(text, fragment) != 1 {
			return fmt.Errorf("qualified node %s missing/duplicated", fragment)
		}
	}
	return nil
}
func (w *contract20BlankWorld) formula() error {
	before := string(w.read.parts[w.part])
	after := string(w.output[w.part])
	if strings.Count(before, `<x:c r="C1"><x:f>1+1</x:f><x:v>2</x:v></x:c>`) != 1 || strings.Count(after, `<x:c r="C1"><x:f>1+1</x:f></x:c>`) != 1 {
		return fmt.Errorf("C1 formula/cache differ")
	}
	return nil
}
func (w *contract20BlankWorld) calculation() error {
	var wb struct {
		Calc []struct {
			Mode  string `xml:"calcMode,attr"`
			Full  string `xml:"fullCalcOnLoad,attr"`
			Force string `xml:"forceFullCalc,attr"`
		} `xml:"calcPr"`
	}
	if err := xml.Unmarshal(w.output["xl/workbook.xml"], &wb); err != nil {
		return err
	}
	if len(wb.Calc) != 1 || wb.Calc[0].Mode != "auto" || wb.Calc[0].Full != "1" || wb.Calc[0].Force != "1" || strings.Count(string(w.output["xl/workbook.xml"]), "<x:calcPr") != 1 {
		return fmt.Errorf("qualified calcPr missing or wrong")
	}
	return nil
}
func (w *contract20BlankWorld) changed() error {
	want := []string{w.part}
	if w.sheet == "Prefixed" {
		want = append(want, "xl/workbook.xml")
	}
	sort.Strings(want)
	got := []string{}
	for part, old := range w.read.parts {
		if !bytes.Equal(old, w.output[part]) {
			got = append(got, part)
		}
	}
	sort.Strings(got)
	if !reflect.DeepEqual(want, got) {
		return fmt.Errorf("changed members %v want %v", got, want)
	}
	return nil
}

// Independent lexical control masks only the named source element spans and
// the explicit ordinary cache / calcPr insertion sites. Other bytes are equal.
func maskBlankCell(source, ref string) (string, error) {
	needle := ` r="` + ref + `"`
	start := strings.Index(source, needle)
	if start < 0 {
		return "", fmt.Errorf("missing %s", ref)
	}
	left := strings.LastIndex(source[:start], "<")
	if left < 0 {
		return "", fmt.Errorf("missing start %s", ref)
	}
	nameEnd := strings.IndexAny(source[left+1:], " \t\r\n/>")
	if nameEnd < 0 {
		return "", fmt.Errorf("bad QName %s", ref)
	}
	name := source[left+1 : left+1+nameEnd]
	tagEnd := strings.Index(source[start:], ">")
	if tagEnd < 0 {
		return "", fmt.Errorf("bad tag %s", ref)
	}
	tagEnd += start + 1
	end := tagEnd
	if source[tagEnd-2] != '/' {
		close := strings.Index(source[tagEnd:], "</"+name+">")
		if close < 0 {
			return "", fmt.Errorf("unclosed %s", ref)
		}
		end = tagEnd + close + len(name) + 3
	}
	return source[:left] + "[SELECTED]" + source[end:], nil
}
func (w *contract20BlankWorld) lexical() error {
	old, newXML := string(w.read.parts[w.part]), string(w.output[w.part])
	var err error
	for _, e := range w.edits {
		old, err = maskBlankCell(old, e.Address)
		if err != nil {
			return err
		}
		newXML, err = maskBlankCell(newXML, e.Address)
		if err != nil {
			return err
		}
	}
	if w.sheet == "Prefixed" {
		const cache = `<x:v>2</x:v>`
		if strings.Count(old, cache) != 1 {
			return fmt.Errorf("unclassifiable source cache")
		}
		old = strings.Replace(old, cache, "[CACHE]", 1)
		marker := `<x:f>1+1</x:f>`
		if strings.Count(newXML, marker) != 1 {
			return fmt.Errorf("formula span changed")
		}
		newXML = strings.Replace(newXML, marker, marker+"[CACHE]", 1)
		prior, post := string(w.read.parts["xl/workbook.xml"]), string(w.output["xl/workbook.xml"])
		start := strings.Index(post, "<x:calcPr")
		if start < 0 {
			return fmt.Errorf("calcPr insertion absent")
		}
		end := strings.Index(post[start:], "/>")
		if end < 0 {
			return fmt.Errorf("calcPr not self closing")
		}
		post = post[:start] + post[start+end+2:]
		if prior != post {
			return fmt.Errorf("unselected workbook bytes changed")
		}
	}
	if old != newXML {
		return fmt.Errorf("unselected worksheet XML bytes changed")
	}
	return nil
}
func (w *contract20BlankWorld) styleCustody() error {
	if !bytes.Equal(w.read.parts["xl/styles.xml"], w.output["xl/styles.xml"]) {
		return fmt.Errorf("style definitions changed")
	}
	if err := w.memberCustody(); err != nil {
		return err
	}
	return w.lexical()
}
func (w *contract20BlankWorld) refusalAndFault() error {
	// Refusal on the selected source; preserving a held unrelated numeric cell
	// is checked where the styled recipe provides one.
	s, err := spreadsheet.OpenEditing(w.read.source, packaging.Limits{})
	if err != nil {
		return err
	}
	var held *spreadsheet.NumberTarget
	if w.sheet == "StyledBlank" {
		held, err = s.FindNumber(w.sheet, "B1")
		if err != nil {
			return err
		}
	}
	err = s.SetContractBlankValues(w.sheet, []spreadsheet.ContractBlankEdit{{Address: "ZZ9", Value: "no"}})
	var refused *packaging.Refusal
	if !errors.As(err, &refused) || refused.Kind != "missing_target" {
		return fmt.Errorf("blank ZZ9 refusal = %v, want missing_target", err)
	}
	checkHeld := func() error {
		if held != nil {
			value, e := s.CurrentNumber(held)
			if e != nil || value != 5 {
				return fmt.Errorf("held StyledBlank!B1 after refusal/fault = %v/%v", value, e)
			}
		}
		return nil
	}
	if err := checkHeld(); err != nil {
		return err
	}
	// The symlink is the actual SaveAs destination, not an unrelated sentinel.
	// Refusal must preserve both its identity and the regular file it targets.
	prior := filepath.Join(w.temp, "existing.xlsx")
	link := filepath.Join(w.temp, "existing-destination.xlsx")
	const sentinel = "existing destination sentinel"
	if err := os.WriteFile(prior, []byte(sentinel), 0600); err != nil {
		return err
	}
	if err := os.Symlink(filepath.Base(prior), link); err != nil {
		return err
	}
	failed := filepath.Join(w.temp, "destination-directory")
	if err := os.Mkdir(failed, 0700); err != nil {
		return err
	}
	// No parent exists for the absent destination: failure must not create it.
	absentParent := filepath.Join(w.temp, "absent-parent")
	absent := filepath.Join(absentParent, "new.xlsx")
	entries := func() ([]string, error) {
		found, e := os.ReadDir(w.temp)
		if e != nil {
			return nil, e
		}
		names := make([]string, 0, len(found))
		for _, item := range found {
			names = append(names, item.Name())
		}
		return names, nil
	}
	before, err := entries()
	if err != nil {
		return err
	}
	for _, destination := range []string{link, absent, failed} {
		if _, err := s.SaveAs(destination); err == nil {
			return fmt.Errorf("failed SaveAs destination accepted: %s", destination)
		}
		after, e := entries()
		if e != nil || !reflect.DeepEqual(before, after) {
			return fmt.Errorf("failed SaveAs changed temporary directory: %v, %v -> %v", e, before, after)
		}
		if err := checkHeld(); err != nil {
			return err
		}
	}
	if target, e := os.Readlink(link); e != nil || target != filepath.Base(prior) {
		return fmt.Errorf("existing SaveAs destination symlink changed: %q/%v", target, e)
	}
	data, err := os.ReadFile(prior)
	if err != nil || string(data) != sentinel {
		return fmt.Errorf("existing destination content changed: %q/%v", data, err)
	}
	if _, err := os.Lstat(absentParent); !os.IsNotExist(err) {
		return fmt.Errorf("absent destination parent created: %v", err)
	}
	if info, err := os.Lstat(failed); err != nil || !info.IsDir() {
		return fmt.Errorf("failed directory destination corrupted: %v", err)
	}
	// A fresh save after every refusal/fault proves the retained session still
	// carries exactly the original payloads; this is not an output hash oracle.
	path := filepath.Join(w.temp, "refused-session.xlsx")
	if _, err = s.SaveAs(path); err != nil {
		return err
	}
	data, err = os.ReadFile(path)
	if err != nil {
		return err
	}
	members, err := blankMemberPayloads(data)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(members, w.read.parts) {
		return fmt.Errorf("refused session changed member payloads")
	}
	if err := checkHeld(); err != nil {
		return err
	}
	return w.inputCustody()
}

func TestContract20BlankSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct{ id, recipe string }{
		{"@id-xlsx-prefixed-namespace-safe-edits", "prefixed-cells"},
		{"@id-xlsx-styled-blank-cell-editable", "styled-blank"},
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
			t.Fatalf("selected blank case %s has %d rows", tc.id, count)
		}
		t.Run(strings.TrimPrefix(tc.id, "@id-"), func(t *testing.T) {
			w := &contract20BlankWorld{temp: t.TempDir()}
			exact := func(sc *godog.ScenarioContext, text string, action func() error) {
				sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
			}
			init := func(sc *godog.ScenarioContext) {
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
					*w = contract20BlankWorld{temp: w.temp}
					return ctx, nil
				})
				if tc.recipe == "prefixed-cells" {
					exact(sc, `the literal contract20 XLSX recipe "prefixed-cells" includes sealed styles indexes 0 1 and 2`, func() error { return w.input(tc.recipe) })
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, w.record)
					exact(sc, `the production value editor writes JSON "Alpha" to Prefixed!A1 and numeric 7 to Prefixed!B1`, w.write)
					exact(sc, `the saved A1 and B1 retain original r and style indexes 1 and 2 and read JSON "Alpha" and numeric 7`, w.values)
					exact(sc, `the saved value nodes are qualified x:is x:t and x:v with the original spreadsheet namespace`, w.qualified)
					exact(sc, `the saved C1 formula remains 1+1 and its x:v cache is absent`, w.formula)
					exact(sc, `the saved workbook has exactly one x:calcPr with calcMode auto fullCalcOnLoad 1 and forceFullCalc 1`, w.calculation)
					exact(sc, `the exact changed member set is JSON ["xl/workbook.xml","xl/worksheets/sheet1.xml"] without additions or removals`, w.changed)
					exact(sc, `only entire original A1 and B1 cell spans, original C1 x:v and the calcPr insertion site may differ; every intervening XML span retains literal bytes`, w.lexical)
					exact(sc, `the original styles.xml and all unpatched original cell attributes and properties retain literal bytes`, w.styleCustody)
				} else {
					exact(sc, `the literal contract20 XLSX recipe "styled-blank" includes sealed styles indexes 0 1 and 2`, func() error { return w.input(tc.recipe) })
					exact(sc, `the original StyledBlank!A1 reads null with style index 1 and B1 reads numeric 5 with style index 2`, w.beforeStyled)
					exact(sc, `the production value editor writes JSON "filled" to StyledBlank!A1`, w.write)
					exact(sc, `the reopened StyledBlank!A1 reads JSON "filled" with original style index 1`, w.values)
					exact(sc, `the reopened StyledBlank!B1 reads numeric 5 with original style index 2`, w.values)
					exact(sc, `the exact changed member set is JSON ["xl/worksheets/sheet1.xml"] without additions or removals`, w.changed)
					exact(sc, `only the selected original A1 cell element span may differ; styles.xml B1 and every other original XML span remain literal`, w.lexical)
					exact(sc, `every original style definition and relationship retains its original payload`, w.styleCustody)
				}
				exact(sc, `the result is saved to a distinct new path and independently parsed and reopened`, w.save)
				exact(sc, `the source fixture, caller archive and operation operands remain unchanged`, w.inputCustody)
				exact(sc, `only the sealed original member and lexical span allowances differ; all unrelated member payloads remain literal`, w.memberCustody)
				exact(sc, `all saved OPC relationships and content types resolve with original identities preserved`, w.relationships)
				exact(sc, `refusals and save faults publish no partial destination and leave the session and held unaffected targets usable`, w.refusalAndFault)
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-blank-" + strings.TrimPrefix(tc.id, "@"), ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: tc.id, Strict: true, Concurrency: 1}}
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
			if dir := os.Getenv("GO_CONTRACT20_BLANK_RECEIPTS_DIR"); dir != "" && !t.Failed() {
				if err := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(tc.id, "@id-")+".cucumber.json"), output.Bytes(), 0600); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
}
