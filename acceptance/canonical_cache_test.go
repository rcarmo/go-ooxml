package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

// Only the exact manifest-backed cross-sheet case is registered here. The
// synthetic CACHE-001 source remains inventoried but is not selected.
func canonicalCacheSteps(sc *godog.ScenarioContext) {
	const assetID = "fixture-8ba5708d5030adf93a4f7e4ae466563a1b66341067a200a6cb9b782484dcb5b1"
	const fixtureID = "cross-sheet-cache.xlsx"
	var fixture mutationFixture
	var sourcePath, destination string
	var source, recorded []byte
	var session *spreadsheet.EditSession
	var effect spreadsheet.CalculationEffect
	var receipt packaging.Receipt
	var changed bool

	clear := func() {
		fixture = mutationFixture{}
		sourcePath, destination = "", ""
		source, recorded = nil, nil
		session = nil
		effect = spreadsheet.CalculationEffect{}
		changed = false
	}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		clear()
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if sourcePath != "" {
			_ = os.RemoveAll(filepath.Dir(sourcePath))
		}
		return ctx, nil
	})
	readDestination := func() ([]byte, map[string][]byte, error) {
		if destination == "" {
			return nil, nil, fmt.Errorf("destination not declared")
		}
		b, err := os.ReadFile(destination)
		if err != nil {
			return nil, nil, err
		}
		members, err := zipPayloads(b)
		return b, members, err
	}
	cell := func(member, address, value, formula string) error {
		_, members, err := readDestination()
		if err != nil {
			return err
		}
		gotValue, gotFormula, err := sharedCell(members[member], address)
		if err != nil {
			return err
		}
		if gotValue != value || gotFormula != formula {
			return fmt.Errorf("%s:%s value/formula %q/%q, want %q/%q", member, address, gotValue, gotFormula, value, formula)
		}
		return nil
	}

	sc.Step(`^fixture "cross-sheet-cache\.xlsx" verified against the fixture manifest$`, func() error {
		contractBytes, err := os.ReadFile(testutil.ReferencePath("contracts/mutation-safety.json"))
		if err != nil {
			return err
		}
		contract, err := parseMutationContract(contractBytes)
		if err != nil {
			return err
		}
		for _, item := range contract.Fixtures {
			if item.ID == fixtureID {
				fixture = item
			}
		}
		if fixture.AssetID != assetID {
			return fmt.Errorf("shared cache fixture identity missing/changed: %q", fixture.AssetID)
		}
		path, err := testutil.LookupFixture(fixture.AssetID)
		if err != nil {
			return err
		}
		data, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		if sha256hex(data) != strings.TrimPrefix(assetID, "fixture-") {
			return fmt.Errorf("manifest-backed input hash changed")
		}
		members, err := zipPayloads(data)
		if err != nil {
			return err
		}
		if len(members) != len(fixture.Members) {
			return fmt.Errorf("fixture member count changed")
		}
		for name, hash := range fixture.Members {
			if sha256hex(members[name]) != hash {
				return fmt.Errorf("fixture member changed: %s", name)
			}
		}
		allowed := map[string]bool{
			"xl/workbook.xml":          true,
			"xl/worksheets/sheet1.xml": true,
			"xl/worksheets/sheet2.xml": true,
		}
		if fixture.Facts["Input!A1"] != float64(1) || fixture.Facts["calcChainPresent"] != false || len(fixture.Allowed) != len(allowed) {
			return fmt.Errorf("unexpected cache contract facts/allowance")
		}
		formulaFact, ok := fixture.Facts["Calc!A1"].(map[string]any)
		if !ok || formulaFact["formula"] != "=Input!A1*2" || formulaFact["cachedValue"] != float64(2) {
			return fmt.Errorf("unexpected Calc!A1 fixture fact")
		}
		for _, name := range fixture.Allowed {
			if !allowed[name] {
				return fmt.Errorf("unexpected allowed member %s", name)
			}
			delete(allowed, name)
		}
		if len(allowed) != 0 {
			return fmt.Errorf("missing allowed members %v", allowed)
		}
		if _, hasChain := members["xl/calcChain.xml"]; hasChain {
			return fmt.Errorf("fixture unexpectedly contains calculation chain")
		}
		if value, formula, err := sharedCell(members["xl/worksheets/sheet1.xml"], "A1"); err != nil || value != "1" || formula != "" {
			return fmt.Errorf("invalid initial Input!A1: %q/%q: %v", value, formula, err)
		}
		if value, formula, err := sharedCell(members["xl/worksheets/sheet2.xml"], "A1"); err != nil || value != "2" || formula != "Input!A1*2" {
			return fmt.Errorf("invalid initial Calc!A1: %q/%q: %v", value, formula, err)
		}
		source = data
		return nil
	})
	sc.Step(`^destination state is "distinct-absent"$`, func() error {
		if len(source) == 0 {
			return fmt.Errorf("source fixture not verified")
		}
		dir, err := os.MkdirTemp("", "shared-cache-canonical-")
		if err != nil {
			return err
		}
		sourcePath = filepath.Join(dir, "source.xlsx")
		destination = filepath.Join(dir, "destination.xlsx")
		if err := os.WriteFile(sourcePath, source, 0600); err != nil {
			return err
		}
		if _, err := os.Stat(destination); !os.IsNotExist(err) {
			return fmt.Errorf("destination was not absent: %v", err)
		}
		return nil
	})
	sc.Step(`^source bytes are recorded$`, func() error {
		data, err := os.ReadFile(sourcePath)
		if err != nil {
			return err
		}
		recorded = bytes.Clone(data)
		return nil
	})
	sc.Step(`^calculation policy is "invalidate-without-recalculation"$`, func() error {
		if len(recorded) == 0 {
			return fmt.Errorf("source not recorded")
		}
		var err error
		session, err = spreadsheet.OpenEditing(recorded, packaging.Limits{MaxSourceBytes: 64 << 20, MaxEntries: 4096, MaxPartBytes: 32 << 20, MaxTotalBytes: 128 << 20})
		return err
	})
	sc.Step(`^committing this batch to the distinct destination:$`, func(table *godog.Table) error {
		if session == nil || len(table.Rows) != 2 || len(table.Rows[0].Cells) != 2 || len(table.Rows[1].Cells) != 2 || table.Rows[0].Cells[0].Value != "target" || table.Rows[0].Cells[1].Value != "value_json" || table.Rows[1].Cells[0].Value != "Input!A1" || table.Rows[1].Cells[1].Value != "10" {
			return fmt.Errorf("unsupported canonical batch table")
		}
		if _, err := os.Stat(destination); !os.IsNotExist(err) {
			return fmt.Errorf("destination no longer absent: %v", err)
		}
		target, err := session.FindNumber("Input", "A1")
		if err != nil {
			return err
		}
		effect, err = session.SetNumberWithInvalidation(target, 10)
		if err != nil {
			return err
		}
		receipt, err = session.SaveAs(destination)
		if err == nil {
			changed = true
		}
		return err
	})
	sc.Step(`^committed change count is 1$`, func() error {
		if !changed || !effect.ValueChanged || len(effect.Invalidated) != 1 || len(receipt.Changes) != 3 {
			return fmt.Errorf("unexpected commit effect %+v", effect)
		}
		return nil
	})
	sc.Step(`^source bytes equal the recorded state$`, func() error {
		actual, err := os.ReadFile(sourcePath)
		if err != nil {
			return err
		}
		if !bytes.Equal(recorded, actual) || !bytes.Equal(recorded, source) {
			return fmt.Errorf("source mutated")
		}
		return nil
	})
	sc.Step(`^the reopened destination Input!A1 is numeric 10$`, func() error {
		if err := cell("xl/worksheets/sheet1.xml", "A1", "10", ""); err != nil {
			return err
		}
		// Open the delivered path independently of the live editing session.
		reopened, err := spreadsheet.Open(destination)
		if err != nil {
			return err
		}
		defer reopened.Close()
		sheet, err := reopened.Sheet("Input")
		if err != nil {
			return err
		}
		input := sheet.Cell("A1")
		if input == nil || input.Value() != float64(10) || input.Formula() != "" {
			return fmt.Errorf("reopened Input!A1 is not numeric 10")
		}
		return nil
	})
	sc.Step(`^the reopened destination Calc!A1 formula is "=Input!A1\*2"$`, func() error {
		reopened, err := spreadsheet.Open(destination)
		if err != nil {
			return err
		}
		defer reopened.Close()
		sheet, err := reopened.Sheet("Calc")
		if err != nil {
			return err
		}
		result := sheet.Cell("A1")
		if result == nil || result.Formula() != "Input!A1*2" {
			return fmt.Errorf("reopened Calc!A1 formula missing")
		}
		return nil
	})
	sc.Step(`^the destination Calc!A1 cached value is absent or empty$`, func() error {
		_, members, err := readDestination()
		if err != nil {
			return err
		}
		value, _, err := sharedCell(members["xl/worksheets/sheet2.xml"], "A1")
		if err != nil || value != "" {
			return fmt.Errorf("reopened stale cache %q: %v", value, err)
		}
		return nil
	})
	sc.Step(`^calculation state is "recalculation-required"$`, func() error {
		if effect.State != "recalculation-required" {
			return fmt.Errorf("effect state %q", effect.State)
		}
		_, members, err := readDestination()
		if err != nil {
			return err
		}
		d, err := losslessxml.Parse(members["xl/workbook.xml"])
		if err != nil {
			return err
		}
		for _, e := range d.Elements() {
			if e.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "calcPr"}) {
				continue
			}
			flags := map[string]string{}
			for _, a := range e.Attributes() {
				flags[a.Name.Local] = a.Value
			}
			if flags["calcMode"] == "auto" && flags["fullCalcOnLoad"] == "1" && flags["forceFullCalc"] == "1" {
				return nil
			}
		}
		return fmt.Errorf("full-recalculation flags missing")
	})
	sc.Step(`^a data-only read never returns the old cached value 2 as current$`, func() error {
		z, err := zip.OpenReader(destination)
		if err != nil {
			return err
		}
		defer z.Close()
		count := 0
		for _, f := range z.File {
			if f.Name != "xl/worksheets/sheet2.xml" {
				continue
			}
			count++
			r, err := f.Open()
			if err != nil {
				return err
			}
			b, readErr := io.ReadAll(r)
			closeErr := r.Close()
			if readErr != nil {
				return readErr
			}
			if closeErr != nil {
				return closeErr
			}
			var worksheet struct {
				Cells []struct {
					Address string `xml:"r,attr"`
					Value   string `xml:"v"`
				} `xml:"sheetData>row>c"`
			}
			if err := xml.Unmarshal(b, &worksheet); err != nil {
				return err
			}
			found := 0
			for _, c := range worksheet.Cells {
				if c.Address == "A1" {
					found++
					if c.Value != "" {
						return fmt.Errorf("data-only Calc!A1 exposes cached %q", c.Value)
					}
				}
			}
			if found != 1 {
				return fmt.Errorf("expected one data-only Calc!A1, got %d", found)
			}
		}
		if count != 1 {
			return fmt.Errorf("expected one Calc worksheet, got %d", count)
		}
		return nil
	})
	sc.Step(`^all destination relationship and content-type references resolve$`, func() error {
		b, _, err := readDestination()
		if err != nil {
			return err
		}
		q, err := packaging.OpenPreserved(b, packaging.Limits{})
		if err != nil {
			return err
		}
		g, err := q.Graph()
		if err != nil {
			return err
		}
		sentinel := false
		for _, e := range g.Edges {
			if e.ResolvedPart == "customXml/preservation-sentinel.xml" {
				sentinel = true
			}
		}
		if !sentinel {
			return fmt.Errorf("preservation sentinel relationship lost")
		}
		return nil
	})
	sc.Step(`^destination member payloads outside the manifest change allowance are byte-identical$`, func() error {
		_, actual, err := readDestination()
		if err != nil {
			return err
		}
		original, err := zipPayloads(recorded)
		if err != nil {
			return err
		}
		if len(original) != len(fixture.Members) || len(actual) != len(original) {
			return fmt.Errorf("unexpected member set")
		}
		allowed := map[string]bool{}
		for _, name := range fixture.Allowed {
			allowed[name] = true
		}
		for name, before := range original {
			after, exists := actual[name]
			if !exists {
				return fmt.Errorf("member removed: %s", name)
			}
			if !allowed[name] && !bytes.Equal(before, after) {
				return fmt.Errorf("preserved member changed: %s", name)
			}
		}
		for name := range actual {
			if _, exists := original[name]; !exists {
				return fmt.Errorf("member added: %s", name)
			}
		}
		return nil
	})
}
