package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

type contract20ReadWorld struct {
	source   []byte
	original []byte
	parts    map[string][]byte
	wb       spreadsheet.Workbook
}

func (w *contract20ReadWorld) input(id string) error {
	b, err := os.ReadFile(testutil.ReferencePath("ledgers", "contract20-recipes.json"))
	if err != nil {
		return err
	}
	h := sha256.Sum256(b)
	if hex.EncodeToString(h[:]) != contract20Artifacts["ledgers/contract20-recipes.json"] {
		return fmt.Errorf("recipe provenance differs")
	}
	var recipes struct {
		XLSX []struct {
			ID      string            `json:"id"`
			Members map[string]string `json:"members"`
		} `json:"xlsx"`
	}
	if err := json.Unmarshal(b, &recipes); err != nil {
		return err
	}
	var members map[string]string
	for _, r := range recipes.XLSX {
		if r.ID == id {
			members = r.Members
			break
		}
	}
	if members == nil {
		return fmt.Errorf("missing literal input %s", id)
	}
	var output bytes.Buffer
	zipper := zip.NewWriter(&output)
	names := make([]string, 0, len(members))
	for name := range members {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		if !filepath.IsLocal(name) || strings.Contains(name, "\\") {
			return fmt.Errorf("unsafe literal member %s", name)
		}
		writer, e := zipper.Create(name)
		if e != nil {
			return e
		}
		if _, e = writer.Write([]byte(members[name])); e != nil {
			return e
		}
	}
	if err := zipper.Close(); err != nil {
		return err
	}
	w.source = output.Bytes()
	w.original = append([]byte(nil), w.source...)
	w.parts = map[string][]byte{}
	reader, err := zip.NewReader(bytes.NewReader(w.source), int64(len(w.source)))
	if err != nil {
		return err
	}
	for _, file := range reader.File {
		if _, seen := w.parts[file.Name]; seen {
			return fmt.Errorf("duplicate literal member %s", file.Name)
		}
		r, e := file.Open()
		if e != nil {
			return e
		}
		data, e := io.ReadAll(r)
		ce := r.Close()
		if e != nil {
			return e
		}
		if ce != nil {
			return ce
		}
		w.parts[file.Name] = data
	}
	return nil
}
func (w *contract20ReadWorld) open() error {
	var err error
	w.wb, err = spreadsheet.OpenReader(bytes.NewReader(w.source), int64(len(w.source)))
	return err
}
func (w *contract20ReadWorld) custody() error {
	if !bytes.Equal(w.source, w.original) {
		return fmt.Errorf("caller archive bytes changed during read")
	}
	reader, err := zip.NewReader(bytes.NewReader(w.source), int64(len(w.source)))
	if err != nil {
		return err
	}
	if len(reader.File) != len(w.parts) {
		return fmt.Errorf("read changed member count")
	}
	for _, file := range reader.File {
		r, e := file.Open()
		if e != nil {
			return e
		}
		data, e := io.ReadAll(r)
		ce := r.Close()
		if e != nil {
			return e
		}
		if ce != nil {
			return ce
		}
		if !bytes.Equal(data, w.parts[file.Name]) {
			return fmt.Errorf("member changed: %s", file.Name)
		}
	}
	return nil
}
func (w *contract20ReadWorld) linked() error {
	if w.wb.SheetCount() != 2 {
		return fmt.Errorf("sheet count %d", w.wb.SheetCount())
	}
	for index, want := range []string{"Alpha", "Beta"} {
		sheet, e := w.wb.Sheet(index)
		if e != nil {
			return e
		}
		if sheet.Name() != want {
			return fmt.Errorf("sheet %d = %s", index, sheet.Name())
		}
	}
	workbook, rels := string(w.parts["xl/workbook.xml"]), string(w.parts["xl/_rels/workbook.xml.rels"])
	for _, text := range []string{`name="Alpha" sheetId="1" r:id="rId2"`, `name="Beta" sheetId="2" r:id="rId1"`} {
		if !strings.Contains(workbook, text) {
			return fmt.Errorf("missing workbook edge %s", text)
		}
	}
	for _, text := range []string{`Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"`, `Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"`} {
		if !strings.Contains(rels, text) {
			return fmt.Errorf("missing relationship %s", text)
		}
	}
	return nil
}
func (w *contract20ReadWorld) values() error {
	for _, item := range []struct {
		name, shared   string
		number         float64
		boolean        bool
		formula, cache string
	}{{"Alpha", "alpha shared", 7, false, "B1+1", "8"}, {"Beta", "beta shared", 42, true, "B1*2", "84"}} {
		sheet, e := w.wb.Sheet(item.name)
		if e != nil {
			return e
		}
		if got := sheet.Cell("A1").String(); got != item.shared {
			return fmt.Errorf("%s A1 %q", item.name, got)
		}
		if got, e := sheet.Cell("B1").Float64(); e != nil || got != item.number {
			return fmt.Errorf("%s B1 %v/%v", item.name, got, e)
		}
		if got, e := sheet.Cell("C1").Bool(); e != nil || got != item.boolean {
			return fmt.Errorf("%s C1 %v/%v", item.name, got, e)
		}
		if got := sheet.Cell("D1").Formula(); got != item.formula {
			return fmt.Errorf("%s D1 formula %q", item.name, got)
		}
		if got := sheet.Cell("D1").String(); got != item.cache {
			return fmt.Errorf("%s D1 cache %q", item.name, got)
		}
	}
	return nil
}
func (w *contract20ReadWorld) phonetics() error {
	sheet, e := w.wb.Sheet("Phonetics")
	if e != nil {
		return e
	}
	for _, item := range []struct{ cell, text string }{{"A1", "Alpha Beta"}, {"B1", "Gamma Delta"}} {
		if got := sheet.Cell(item.cell).String(); got != item.text {
			return fmt.Errorf("%s text %q, want %q", item.cell, got, item.text)
		}
	}
	return nil
}
func (w *contract20ReadWorld) phoneticCustody() error {
	if err := w.custody(); err != nil {
		return err
	}
	for _, name := range []string{"xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"} {
		text := string(w.parts[name])
		if !strings.Contains(text, `<rPh sb="0" eb="5"><t>Guide</t></rPh>`) || !strings.Contains(text, `xml:space="preserve"`) {
			return fmt.Errorf("phonetic source XML changed %s", name)
		}
	}
	return nil
}

// These are the only two currently bound Contract20 cases. Missing bindings for
// the remaining 20 are deliberately not converted into passing placeholders.
func TestContract20ReadSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	const linkedID = "@id-xlsx-read-rel-linked-shared-strings"
	const phoneticID = "@id-xlsx-phonetic-guides-excluded"
	for _, id := range []string{linkedID, phoneticID} {
		var key caseID
		var want expectedCase
		count := 0
		for candidate, record := range inventory {
			if candidate.ID == id {
				key, want = candidate, record
				count++
			}
		}
		if count != 1 {
			t.Fatalf("selected read case %s has %d rows", id, count)
		}
		t.Run(strings.TrimPrefix(id, "@id-"), func(t *testing.T) {
			state := &contract20ReadWorld{}
			exact := func(sc *godog.ScenarioContext, text string, action func() error) {
				sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
			}
			init := func(sc *godog.ScenarioContext) {
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
					*state = contract20ReadWorld{}
					return ctx, nil
				})
				sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
					if state.wb != nil {
						_ = state.wb.Close()
					}
					return ctx, nil
				})
				if id == linkedID {
					exact(sc, `the literal contract20 XLSX recipe "rel-values" is packed without workbook APIs`, func() error { return state.input("rel-values") })
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, func() error {
						if len(state.parts) == 0 {
							return fmt.Errorf("no input recorded")
						}
						return nil
					})
					exact(sc, `the production read API inspects worksheets in workbook sheet order and obtains literal cell and formula/cache values`, state.open)
					exact(sc, `the ordered sheet identities equal JSON [{"name":"Alpha","part":"xl/worksheets/sheet2.xml"},{"name":"Beta","part":"xl/worksheets/sheet1.xml"}]`, state.linked)
					exact(sc, `the exact read records equal JSON {"Alpha":{"A1":{"kind":"string","value":"alpha shared"},"B1":{"kind":"number","value":7},"C1":{"kind":"boolean","value":false},"D1":{"kind":"formula","formula":"B1+1","cached":8}},"Beta":{"A1":{"kind":"string","value":"beta shared"},"B1":{"kind":"number","value":42},"C1":{"kind":"boolean","value":true},"D1":{"kind":"formula","formula":"B1*2","cached":84}}}`, state.values)
					exact(sc, `all read calls leave the archive, every member payload and relationship unchanged and create no new members`, state.custody)
				} else {
					exact(sc, `the literal contract20 XLSX recipe "phonetic-strings" is packed without workbook APIs`, func() error { return state.input("phonetic-strings") })
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, func() error {
						if len(state.parts) == 0 {
							return fmt.Errorf("no input recorded")
						}
						return nil
					})
					exact(sc, `the production reader reads Phonetics!A1 and Phonetics!B1`, state.open)
					exact(sc, `the exact string values equal JSON {"A1":"Alpha Beta","B1":"Gamma Delta"}`, state.phonetics)
					exact(sc, `no returned visible string includes JSON "Guide"`, state.phonetics)
					exact(sc, `every original rich run and rPh XML span, member payload and caller archive remains unchanged`, state.phoneticCustody)
				}
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-read-" + strings.TrimPrefix(id, "@"), ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: id, Strict: true, Concurrency: 1}}
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
			if dir := os.Getenv("GO_CONTRACT20_XLSX_READ_RECEIPTS_DIR"); dir != "" && !t.Failed() {
				if err := os.MkdirAll(dir, 0o755); err != nil {
					t.Fatal(err)
				}
				if err := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(id, "@id-")+".cucumber.json"), data, 0o644); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
}
