package acceptance

import (
	"archive/zip"
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
	"sort"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

type contract20CacheWorld struct {
	read        contract20ReadWorld
	session     *spreadsheet.EditSession
	held        *spreadsheet.NumberTarget
	effect      spreadsheet.CalculationEffect
	result      error
	recipe, top string
	temp        string
	saved       []byte
	parts       map[string][]byte
	operand     float64
}

func (w *contract20CacheWorld) input(recipe string) error {
	w.recipe = recipe
	if err := w.read.input(recipe); err != nil {
		return err
	}
	w.operand = 10
	var err error
	w.session, err = spreadsheet.OpenEditing(w.read.source, packaging.Limits{})
	if err != nil {
		return err
	}
	w.held, err = w.session.FindNumber("Model", "A1")
	return err
}
func (w *contract20CacheWorld) original() error {
	if w.held == nil {
		return fmt.Errorf("no held input")
	}
	v, err := w.session.CurrentNumber(w.held)
	if err != nil || v != 1 {
		return fmt.Errorf("Model!A1 initial %v/%v", v, err)
	}
	if !bytes.Equal(w.read.original, w.read.source) {
		return fmt.Errorf("original caller archive already changed")
	}
	return nil
}
func (w *contract20CacheWorld) attributed() error {
	if err := w.original(); err != nil {
		return err
	}
	sheet := string(w.read.parts["xl/worksheets/sheet2.xml"])
	if w.recipe == "input-array" {
		w.top = "array"
	} else {
		w.top = "dataTable"
	}
	var parsed struct {
		Rows []struct {
			Cells []struct {
				Address string `xml:"r,attr"`
				Formula struct {
					Kind  string `xml:"t,attr"`
					Range string `xml:"ref,attr"`
					Body  string `xml:",chardata"`
				} `xml:"f"`
				Value *string `xml:"v"`
			} `xml:"c"`
		} `xml:"sheetData>row"`
	}
	if err := xml.Unmarshal([]byte(sheet), &parsed); err != nil {
		return err
	}
	seen := map[string]bool{}
	for _, row := range parsed.Rows {
		for _, cell := range row.Cells {
			switch cell.Address {
			case "A1":
				if seen["A1"] || cell.Formula.Kind != w.top || cell.Formula.Range != "A1:A2" || cell.Formula.Body != "Model!A1*{1;2}" || cell.Value == nil || *cell.Value != "1" {
					return fmt.Errorf("attributed anchor/cache mismatch: %+v", cell)
				}
				seen["A1"] = true
			case "A2":
				if seen["A2"] || cell.Value == nil || *cell.Value != "2" {
					return fmt.Errorf("attributed follower/cache mismatch: %+v", cell)
				}
				seen["A2"] = true
			}
		}
	}
	if len(seen) != 2 {
		return fmt.Errorf("attributed result range incomplete %v", seen)
	}
	return nil
}
func (w *contract20CacheWorld) record() error { return w.original() }
func (w *contract20CacheWorld) inert() error {
	graph, err := w.sessionPackageGraph()
	if err != nil {
		return err
	}
	required := map[string]string{"cache1": "xl/charts/cache-boundary.xml", "cache2": "xl/externalLinks/cache-boundary.xml"}
	found := map[string]bool{}
	wantType := map[string]string{"cache1": "urn:contract20:opaque/chart", "cache2": "urn:contract20:opaque/external"}
	wantContent := map[string]string{"cache1": packaging.ContentTypeChart, "cache2": "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"}
	for _, edge := range graph.Edges {
		if part, ok := required[edge.ID]; ok {
			if edge.Source != "" || edge.ResolvedPart != part || edge.Target != part || edge.External || edge.Type != wantType[edge.ID] || found[edge.ID] {
				return fmt.Errorf("root opaque edge topology differs: %+v", edge)
			}
			found[edge.ID] = true
		}
		for _, part := range required {
			if edge.Source == part || edge.ResolvedPart == part && edge.Source != "" {
				return fmt.Errorf("opaque cache has semantic owner/outgoing edge %+v", edge)
			}
		}
	}
	if len(found) != 2 {
		return fmt.Errorf("inert root edges absent: %v", found)
	}
	parts := map[string]string{}
	for _, part := range graph.Parts {
		parts[part.Name] = part.ContentType
	}
	ct := string(w.read.parts["[Content_Types].xml"])
	for id, part := range required {
		if len(w.read.parts[part]) == 0 || parts[part] != wantContent[id] || strings.Count(ct, `PartName="/`+part+`" ContentType="`+wantContent[id]+`"`) != 1 {
			return fmt.Errorf("opaque part/content override differs: %s", part)
		}
	}
	return nil
}
func (w *contract20CacheWorld) sessionPackageGraph() (packaging.Graph, error) {
	pkg, err := packaging.OpenPreserved(w.read.source, packaging.Limits{})
	if err != nil {
		return packaging.Graph{}, err
	}
	return pkg.Graph()
}
func (w *contract20CacheWorld) write() error {
	if w.recipe == "opaque-caches" {
		w.effect, w.result = w.session.SetNumberWithSealedOpaqueCaches(w.held, 10)
	} else {
		w.effect, w.result = w.session.SetNumberWithAllCachesInvalidated(w.held, 10)
	}
	if w.recipe == "input-array" || w.recipe == "input-dataTable" {
		return nil
	}
	if w.result != nil {
		return w.result
	}
	if !w.effect.ValueChanged || w.effect.State != "recalculation-required" || !reflect.DeepEqual(w.effect.Invalidated, []string{"Model!B1", "Model!B2", "Summary!A1", "Summary!B1"}) {
		return fmt.Errorf("invalid effect %+v", w.effect)
	}
	return nil
}
func (w *contract20CacheWorld) save() error {
	path := filepath.Join(w.temp, "changed.xlsx")
	if _, err := w.session.SaveAs(path); err != nil {
		return err
	}
	b, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	w.saved = b
	w.parts, err = blankMemberPayloads(b)
	if err != nil {
		return err
	}
	wb, err := spreadsheet.OpenReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return err
	}
	defer wb.Close()
	_, err = wb.Sheet("Model")
	if err != nil {
		return err
	}
	_, err = wb.Sheet("Summary")
	return err
}
func (w *contract20CacheWorld) custody() error {
	if !bytes.Equal(w.read.original, w.read.source) || w.operand != 10 {
		return fmt.Errorf("original caller archive or operand changed")
	}
	members, err := blankMemberPayloads(w.read.source)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(members, w.read.parts) {
		return fmt.Errorf("input member payloads changed")
	}
	return nil
}
func (w *contract20CacheWorld) memberCustody() error {
	if len(w.parts) != len(w.read.parts) {
		return fmt.Errorf("member count drift")
	}
	allowed := map[string]bool{"xl/workbook.xml": true, "xl/worksheets/sheet1.xml": true, "xl/worksheets/sheet2.xml": true}
	for part, old := range w.read.parts {
		now, ok := w.parts[part]
		if !ok {
			return fmt.Errorf("missing member %s", part)
		}
		if !allowed[part] && !bytes.Equal(old, now) {
			return fmt.Errorf("unselected member %s changed", part)
		}
	}
	return contract20CacheLexicalCustody(w.read.parts, w.parts)
}
func (w *contract20CacheWorld) relationships() error {
	source, err := packaging.OpenPreserved(w.read.original, packaging.Limits{})
	if err != nil {
		return err
	}
	saved, err := packaging.OpenPreserved(w.saved, packaging.Limits{})
	if err != nil {
		return err
	}
	left, err := source.Graph()
	if err != nil {
		return err
	}
	right, err := saved.Graph()
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(left, right) {
		return fmt.Errorf("saved OPC graph differs")
	}
	for _, part := range []string{"[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels"} {
		if !bytes.Equal(w.read.parts[part], w.parts[part]) {
			return fmt.Errorf("registry %s changed", part)
		}
	}
	return nil
}
func (w *contract20CacheWorld) saveFault() error {
	// Exercise the actual changed session. A held target on the changed input
	// is consumed by the successful edit, so a second independent session
	// tests the refusal/held-target branch of this same predicate.
	if err := contract20SaveFaults(filepath.Join(w.temp, "changed-input-faults"), func(path string) error {
		_, saveErr := w.session.SaveAs(path)
		return saveErr
	}); err != nil {
		return err
	}
	post := filepath.Join(w.temp, "after-fault-changed.xlsx")
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
	if !reflect.DeepEqual(parts, w.parts) {
		return fmt.Errorf("changed session payloads changed after failed saves")
	}
	s, err := spreadsheet.OpenEditing(w.read.source, packaging.Limits{})
	if err != nil {
		return err
	}
	held, err := s.FindNumber("Model", "A1")
	if err != nil {
		return err
	}
	if _, err = s.SetNumberWithAllCachesInvalidated(held, 1); err != nil {
		return err
	}
	if err := contract20SaveFaults(filepath.Join(w.temp, "held-input-faults"), func(path string) error {
		_, saveErr := s.SaveAs(path)
		return saveErr
	}); err != nil {
		return err
	}
	if v, err := s.CurrentNumber(held); err != nil || v != 1 {
		return fmt.Errorf("held after failed save %v/%v", v, err)
	}
	post = filepath.Join(w.temp, "after-fault-noop.xlsx")
	if _, err = s.SaveAs(post); err != nil {
		return err
	}
	data, err = os.ReadFile(post)
	if err != nil {
		return err
	}
	parts, err = blankMemberPayloads(data)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.read.parts) {
		return fmt.Errorf("no-op session member payloads changed after save faults")
	}
	return w.custody()
}
func (w *contract20CacheWorld) values() error {
	wb, err := spreadsheet.OpenReader(bytes.NewReader(w.saved), int64(len(w.saved)))
	if err != nil {
		return err
	}
	defer wb.Close()
	model, err := wb.Sheet("Model")
	if err != nil {
		return err
	}
	if v, e := model.Cell("A1").Float64(); e != nil || v != 10 {
		return fmt.Errorf("reopened Model!A1 %v/%v", v, e)
	}
	if v, e := model.Cell("A2").Float64(); e != nil || v != 2 {
		return fmt.Errorf("reopened Model!A2 %v/%v", v, e)
	}
	return nil
}

var contract20CacheFormulas = map[string]map[string]string{
	"xl/worksheets/sheet1.xml": {"B1": "A1+A2", "B2": "B1*2"},
	"xl/worksheets/sheet2.xml": {"A1": "Model!B1*3", "B1": "40+2"},
}

func (w *contract20CacheWorld) formulas() error {
	for part, want := range contract20CacheFormulas {
		var sheet struct {
			Rows []struct {
				Cells []struct {
					R string  `xml:"r,attr"`
					F string  `xml:"f"`
					V *string `xml:"v"`
				} `xml:"c"`
			} `xml:"sheetData>row"`
		}
		if err := xml.Unmarshal(w.parts[part], &sheet); err != nil {
			return err
		}
		seen := map[string]bool{}
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if expected, ok := want[cell.R]; ok {
					if seen[cell.R] || cell.F != expected || cell.V != nil {
						return fmt.Errorf("%s!%s body/cache = %q/%v", part, cell.R, cell.F, cell.V)
					}
					seen[cell.R] = true
				}
			}
		}
		if len(seen) != len(want) {
			return fmt.Errorf("missing formula cells %s: %v", part, seen)
		}
	}
	return nil
}
func (w *contract20CacheWorld) noCache() error {
	if err := w.formulas(); err != nil {
		return err
	}
	// The retained production data-only reader must return nil for every
	// formula cache; XML absence alone is insufficient to prove this step.
	for _, item := range []struct{ sheet, address string }{{"Model", "B1"}, {"Model", "B2"}, {"Summary", "A1"}, {"Summary", "B1"}} {
		value, err := w.session.ReadFormulaCache(item.sheet, item.address)
		if err != nil || value != nil {
			return fmt.Errorf("data-only %s!%s = %v/%v; want null", item.sheet, item.address, value, err)
		}
	}
	for _, part := range []string{"xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"} {
		for _, cache := range []string{"<v>3</v>", "<v>6</v>", "<v>9</v>", "<v>42</v>"} {
			if strings.Contains(string(w.parts[part]), cache) {
				return fmt.Errorf("old cached value %s survived in %s", cache, part)
			}
		}
	}
	return nil
}
func (w *contract20CacheWorld) calc() error {
	var workbook struct {
		Calc []struct {
			Mode  string `xml:"calcMode,attr"`
			Full  string `xml:"fullCalcOnLoad,attr"`
			Force string `xml:"forceFullCalc,attr"`
		} `xml:"calcPr"`
	}
	if err := xml.Unmarshal(w.parts["xl/workbook.xml"], &workbook); err != nil {
		return err
	}
	if len(workbook.Calc) != 1 || workbook.Calc[0].Mode != "auto" || workbook.Calc[0].Full != "1" || workbook.Calc[0].Force != "1" {
		return fmt.Errorf("calcPr flags incorrect")
	}
	if strings.Count(string(w.parts["xl/workbook.xml"]), "<calcPr") != 1 {
		return fmt.Errorf("calcPr missing or not qualified by source namespace")
	}
	return nil
}
func (w *contract20CacheWorld) changed() error {
	want := []string{"xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"}
	got := []string{}
	for part, old := range w.read.parts {
		if !bytes.Equal(old, w.parts[part]) {
			got = append(got, part)
		}
	}
	sort.Strings(got)
	if !reflect.DeepEqual(got, want) {
		return fmt.Errorf("changed member set %v", got)
	}
	return nil
}
func (w *contract20CacheWorld) noOp() error {
	s, err := spreadsheet.OpenEditing(w.read.original, packaging.Limits{})
	if err != nil {
		return err
	}
	held, err := s.FindNumber("Model", "A1")
	if err != nil {
		return err
	}
	var effect spreadsheet.CalculationEffect
	if w.recipe == "opaque-caches" {
		effect, err = s.SetNumberWithSealedOpaqueCaches(held, 1)
	} else {
		effect, err = s.SetNumberWithAllCachesInvalidated(held, 1)
	}
	if err != nil || effect.ValueChanged || len(effect.Invalidated) != 0 {
		return fmt.Errorf("no-op effect %+v/%v", effect, err)
	}
	path := filepath.Join(w.temp, "noop.xlsx")
	if _, err = s.SaveAs(path); err != nil {
		return err
	}
	out, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	parts, err := blankMemberPayloads(out)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(parts, w.read.parts) {
		return fmt.Errorf("no-op changed a source member/cache/flag")
	}
	for part, formulas := range contract20CacheFormulas {
		var sheet struct {
			Rows []struct {
				Cells []struct {
					Address string  `xml:"r,attr"`
					Formula string  `xml:"f"`
					Value   *string `xml:"v"`
				} `xml:"c"`
			} `xml:"sheetData>row"`
		}
		if err := xml.Unmarshal(parts[part], &sheet); err != nil {
			return err
		}
		seen := map[string]bool{}
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if expected, ok := formulas[cell.Address]; ok {
					old := map[string]string{"xl/worksheets/sheet1.xml/B1": "3", "xl/worksheets/sheet1.xml/B2": "6", "xl/worksheets/sheet2.xml/A1": "9", "xl/worksheets/sheet2.xml/B1": "42"}[part+"/"+cell.Address]
					if seen[cell.Address] || cell.Formula != expected || cell.Value == nil || *cell.Value != old {
						return fmt.Errorf("no-op cache/body changed %s!%s", part, cell.Address)
					}
					seen[cell.Address] = true
				}
			}
		}
		if len(seen) != len(formulas) {
			return fmt.Errorf("no-op formula cells missing %s", part)
		}
	}
	if strings.Count(string(parts["xl/workbook.xml"]), `<calcPr calcId="124519"/>`) != 1 {
		return fmt.Errorf("no-op changed original calcPr flags")
	}
	for _, item := range []struct {
		sheet, address string
		value          float64
	}{{"Model", "B1", 3}, {"Model", "B2", 6}, {"Summary", "A1", 9}, {"Summary", "B1", 42}} {
		v, e := s.ReadFormulaCache(item.sheet, item.address)
		if e != nil || v == nil || *v != item.value {
			return fmt.Errorf("no-op data-only %s!%s %v/%v", item.sheet, item.address, v, e)
		}
	}
	return nil
}
func (w *contract20CacheWorld) refusal() error {
	var refused *packaging.Refusal
	if !errors.As(w.result, &refused) || refused.Kind != "xlsx-cache-topology-unsupported" || w.effect.ValueChanged {
		return fmt.Errorf("attributed topology = %+v/%v", w.effect, w.result)
	}
	return nil
}
func (w *contract20CacheWorld) attributedCustody() error {
	if err := w.custody(); err != nil {
		return err
	}
	v, err := w.session.CurrentNumber(w.held)
	if err != nil || v != 1 {
		return fmt.Errorf("held after refusal = %v/%v", v, err)
	}
	return nil
}
func (w *contract20CacheWorld) refusedSave() error {
	path := filepath.Join(w.temp, "refused.xlsx")
	if _, err := w.session.SaveAs(path); err != nil {
		return err
	}
	out, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	members, err := blankMemberPayloads(out)
	if err != nil {
		return err
	}
	if !reflect.DeepEqual(members, w.read.parts) {
		return fmt.Errorf("refused save changed any member")
	}
	return nil
}
func (w *contract20CacheWorld) refusedDestination() error {
	if err := w.custody(); err != nil {
		return err
	}
	if err := contract20SaveFaults(w.temp, func(path string) error {
		_, saveErr := w.session.SaveAs(path)
		return saveErr
	}); err != nil {
		return err
	}
	if v, err := w.session.CurrentNumber(w.held); err != nil || v != 1 {
		return fmt.Errorf("held after refused save fault %v/%v", v, err)
	}
	post := filepath.Join(w.temp, "after-fault-refused.xlsx")
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
	if !reflect.DeepEqual(parts, w.read.parts) {
		return fmt.Errorf("refused session member payloads changed after save faults")
	}
	return nil
}
func (w *contract20CacheWorld) refusedNoOp() error {
	effect, err := w.session.SetNumberWithAllCachesInvalidated(w.held, 1)
	if err != nil || effect.ValueChanged || len(effect.Invalidated) != 0 {
		return fmt.Errorf("repeated no-op %+v/%v", effect, err)
	}
	if err := w.attributedCustody(); err != nil {
		return err
	}
	part := w.read.parts["xl/worksheets/sheet2.xml"]
	if w.top == "" || !strings.Contains(string(part), `<f t="`+w.top+`" ref="A1:A2">`) || !strings.Contains(string(part), `<c r="A2"><v>2</v></c>`) {
		return fmt.Errorf("refused no-op lost attributed range/cache")
	}
	return nil
}
func (w *contract20CacheWorld) opaquePayloads() error {
	for _, part := range []string{"xl/charts/cache-boundary.xml", "xl/externalLinks/cache-boundary.xml"} {
		old, ok := w.read.parts[part]
		if !ok || !bytes.Equal(old, w.parts[part]) || !strings.Contains(string(old), ">1<") {
			return fmt.Errorf("opaque payload changed or missing literal one: %s", part)
		}
	}
	return nil
}
func contract20PackMembers(members map[string][]byte) ([]byte, error) {
	var output bytes.Buffer
	archive := zip.NewWriter(&output)
	parts := make([]string, 0, len(members))
	for part := range members {
		parts = append(parts, part)
	}
	sort.Strings(parts)
	for _, part := range parts {
		writer, err := archive.Create(part)
		if err != nil {
			return nil, err
		}
		if _, err = writer.Write(members[part]); err != nil {
			return nil, err
		}
	}
	if err := archive.Close(); err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}
func (w *contract20CacheWorld) opaqueVariant() error {
	// Either semantic owner, chart or external, defeats inert-cache admission.
	for _, profile := range []struct{ kind, relation, target string }{
		{"chart", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "charts/cache-boundary.xml"},
		{"external", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink", "externalLinks/cache-boundary.xml"},
	} {
		source, err := blankMemberPayloads(w.read.original)
		if err != nil {
			return err
		}
		rels := string(source["xl/_rels/workbook.xml.rels"])
		if strings.Count(rels, "</Relationships>") != 1 {
			return fmt.Errorf("unknown relationship registry")
		}
		rels = strings.Replace(rels, "</Relationships>", `<Relationship Id="semanticOwner" Type="`+profile.relation+`" Target="`+profile.target+`"/></Relationships>`, 1)
		source["xl/_rels/workbook.xml.rels"] = []byte(rels)
		variant, err := contract20PackMembers(source)
		if err != nil {
			return err
		}
		s, err := spreadsheet.OpenEditing(variant, packaging.Limits{})
		if err != nil {
			return err
		}
		held, err := s.FindNumber("Model", "A1")
		if err != nil {
			return err
		}
		effect, err := s.SetNumberWithSealedOpaqueCaches(held, 10)
		var refused *packaging.Refusal
		if !errors.As(err, &refused) || refused.Kind != "xlsx-opaque-cache-scope-unsupported" || effect.ValueChanged {
			return fmt.Errorf("%s semantic owner accepted: %+v/%v", profile.kind, effect, err)
		}
		if v, e := s.CurrentNumber(held); e != nil || v != 1 {
			return fmt.Errorf("held after %s semantic refusal %v/%v", profile.kind, v, e)
		}
		path := filepath.Join(w.temp, "opaque-"+profile.kind+"-refused.xlsx")
		if _, err = s.SaveAs(path); err != nil {
			return err
		}
		saved, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		parts, err := blankMemberPayloads(saved)
		if err != nil {
			return err
		}
		if !reflect.DeepEqual(source, parts) {
			return fmt.Errorf("%s semantic variant changed saved members", profile.kind)
		}
		before, err := blankMemberPayloads(variant)
		if err != nil {
			return err
		}
		if !reflect.DeepEqual(source, before) {
			return fmt.Errorf("%s semantic variant changed caller bytes", profile.kind)
		}
	}
	return nil
}

func TestContract20CacheSelected(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	inventory, err := contract20Inventory()
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct{ id, recipe string }{
		{"@id-xlsx-clear-cross-sheet-caches", "cross-caches"},
		{"@id-xlsx-array-input-refusal", "input-array"},
		{"@id-xlsx-cache-scope-opaque-parts", "opaque-caches"},
	} {
		var key caseID
		want := map[caseID]expectedCase{}
		for k, v := range inventory {
			if k.ID == tc.id {
				key = k
				want[k] = v
			}
		}
		expected := 1
		if tc.id == "@id-xlsx-array-input-refusal" {
			expected = 2
		}
		if len(want) != expected {
			t.Fatalf("selected cache case %s has %d rows, want %d", tc.id, len(want), expected)
		}
		t.Run(strings.TrimPrefix(tc.id, "@id-")+"/"+tc.recipe, func(t *testing.T) {
			state := &contract20CacheWorld{temp: t.TempDir()}
			exact := func(sc *godog.ScenarioContext, text string, action func() error) {
				sc.Step("^"+regexp.QuoteMeta(text)+"$", action)
			}
			init := func(sc *godog.ScenarioContext) {
				sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
					*state = contract20CacheWorld{temp: t.TempDir()}
					return ctx, nil
				})
				if tc.recipe == "cross-caches" {
					exact(sc, `the literal contract20 XLSX recipe "cross-caches" is packed without workbook APIs`, func() error { return state.input(tc.recipe) })
					exact(sc, `Model!A1 is numeric 1 and the original source and all member payloads are recorded`, state.original)
					exact(sc, `the production value editor writes numeric 10 to Model!A1 under the invalidate-all-ordinary-caches profile`, state.write)
					exact(sc, `the saved Model!A1 is numeric 10 and Model!A2 remains numeric 2`, state.values)
					exact(sc, `the four saved formula bodies equal JSON {"Model!B1":"A1+A2","Model!B2":"B1*2","Summary!A1":"Model!B1*3","Summary!B1":"40+2"}`, state.formulas)
					exact(sc, `all four formula cells have no value cache and data-only reads return null rather than original caches JSON [3,6,9,42]`, state.noCache)
					exact(sc, `the saved workbook has exactly one qualified calcPr with calcMode auto fullCalcOnLoad 1 and forceFullCalc 1`, state.calc)
					exact(sc, `the exact changed member set is JSON ["xl/workbook.xml","xl/worksheets/sheet1.xml","xl/worksheets/sheet2.xml"] without additions or removals`, state.changed)
					exact(sc, `an independent numeric-1 no-op on the original input retains every original member and all four caches and calculation flags`, state.noOp)
				} else if tc.recipe == "opaque-caches" {
					exact(sc, `the literal contract20 XLSX recipe "opaque-caches" is packed without workbook APIs`, func() error { return state.input(tc.recipe) })
					exact(sc, `the exact inert root edges cache1 and cache2 and typed content overrides match the sealed recipe and no semantic chart or external owner edge exists`, state.inert)
					exact(sc, `the source member payloads, relationships and caller archive are recorded`, state.record)
					exact(sc, `the production value editor explicitly opts into sealed-opaque-retention and writes numeric 10 to Model!A1`, state.write)
					exact(sc, `all four ordinary formula bodies remain unchanged and all their cached values are absent`, state.noCache)
					exact(sc, `the saved workbook has exactly one qualified calcPr with calcMode auto fullCalcOnLoad 1 and forceFullCalc 1`, state.calc)
					exact(sc, `both opaque cache payloads retain literal numeric 1 and every original byte`, state.opaquePayloads)
					exact(sc, `the exact changed member set is JSON ["xl/workbook.xml","xl/worksheets/sheet1.xml","xl/worksheets/sheet2.xml"] without additions or removals`, state.changed)
					exact(sc, `an independent variant with an extra semantic chart or external owner relationship refuses before writing any member or destination`, state.opaqueVariant)
				} else {
					for _, recipe := range []string{"input-array", "input-dataTable"} {
						recipe := recipe
						exact(sc, fmt.Sprintf(`the literal contract20 XLSX recipe "%s" is packed without workbook APIs`, recipe), func() error { return state.input(recipe) })
					}
					exact(sc, `the original Model!A1 is numeric 1 and Summary!A1:A2 cached result values are JSON [1,2]`, state.attributed)
					exact(sc, `the source and prior destination bytes, every member and held input-cell target are recorded`, state.record)
					exact(sc, `the production value editor attempts numeric 10 at Model!A1 under the invalidate-all-ordinary-caches profile`, state.write)
					exact(sc, `the operation returns exactly typed refusal category "xlsx-cache-topology-unsupported" before any expression evaluation and no edited result`, state.refusal)
					exact(sc, `all attributed formula metadata, follower cells, cached values and every member payload remain byte-identical`, state.attributedCustody)
					exact(sc, `saving the refused session preserves every original member without additions or removals`, state.refusedSave)
					exact(sc, `source caller operands and prior destination remain unchanged with no partial output`, state.refusedDestination)
					exact(sc, `a repeated numeric-1 no-op observation leaves the refused input and caches unchanged`, state.refusedNoOp)
				}
				if tc.recipe == "cross-caches" || tc.recipe == "opaque-caches" {
					exact(sc, `the result is saved to a distinct new path and independently parsed and reopened`, state.save)
					exact(sc, `the source fixture, caller archive and operation operands remain unchanged`, state.custody)
					exact(sc, `only the sealed original member and lexical span allowances differ; all unrelated member payloads remain literal`, state.memberCustody)
					exact(sc, `all saved OPC relationships and content types resolve with original identities preserved`, state.relationships)
					exact(sc, `refusals and save faults publish no partial destination and leave the session and held unaffected targets usable`, state.saveFault)
				}
			}
			var output bytes.Buffer
			suite := godog.TestSuite{Name: "go-contract20-cache-" + tc.recipe, ScenarioInitializer: init, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{key.File}, Tags: tc.id, Strict: true, Concurrency: 1}}
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
			if err := reconcile(want, data); err != nil {
				t.Error(err)
			}
			if dir := os.Getenv("GO_CONTRACT20_XLSX_RECEIPTS_DIR"); dir != "" && !t.Failed() {
				if e := os.WriteFile(filepath.Join(dir, strings.TrimPrefix(tc.id, "@id-")+".cucumber.json"), output.Bytes(), 0600); e != nil {
					t.Fatal(e)
				}
			}
		})
	}
}
