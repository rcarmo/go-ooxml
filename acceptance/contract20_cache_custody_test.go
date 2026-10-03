package acceptance

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/spreadsheet"
)

type contract20LexicalRemoval struct {
	start, end  int
	replacement string
}

// contract20MaskSpans removes only proved disjoint source ranges. The XML
// parser supplies offsets, but the comparison itself is on original bytes.
func contract20MaskSpans(source []byte, spans []contract20LexicalRemoval) (string, error) {
	var out bytes.Buffer
	last := 0
	for _, s := range spans {
		if s.start < last || s.end < s.start || s.end > len(source) {
			return "", fmt.Errorf("overlapping or invalid lexical span %d..%d", s.start, s.end)
		}
		out.Write(source[last:s.start])
		out.WriteString(s.replacement)
		last = s.end
	}
	out.Write(source[last:])
	return out.String(), nil
}

func contract20SheetLexicalMask(payload []byte, sheet string, saved bool) (string, error) {
	doc, err := losslessxml.Parse(payload)
	if err != nil {
		return "", err
	}
	es := doc.Elements()
	if len(es) == 0 || es[0].Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "worksheet"}) {
		return "", fmt.Errorf("worksheet root differs")
	}
	input := map[string]string{}
	formulas := map[string]struct{ body, cache string }{}
	if sheet == "Model" {
		input["A1"] = "1"
		if saved {
			input["A1"] = "10"
		}
		formulas["B1"] = struct{ body, cache string }{"A1+A2", "3"}
		formulas["B2"] = struct{ body, cache string }{"B1*2", "6"}
	} else if sheet == "Summary" {
		formulas["A1"] = struct{ body, cache string }{"Model!B1*3", "9"}
		formulas["B1"] = struct{ body, cache string }{"40+2", "42"}
	} else {
		return "", fmt.Errorf("unsupported sheet %q", sheet)
	}
	spans := []contract20LexicalRemoval{}
	seen := map[string]bool{}
	for _, cell := range es {
		if cell.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "c"}) {
			continue
		}
		address := ""
		for _, a := range cell.Attributes() {
			if a.Name.Space == "" && a.Name.Local == "r" {
				address = a.Value
			}
		}
		_, isInput := input[address]
		formulaWant, isFormula := formulas[address]
		if !isInput && !isFormula {
			continue
		}
		if seen[address] {
			return "", fmt.Errorf("duplicate selected %s!%s", sheet, address)
		}
		seen[address] = true
		row, ok := cell.Parent()
		if !ok || row.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "row"}) {
			return "", fmt.Errorf("cell row owner differs")
		}
		data, ok := row.Parent()
		if !ok || data.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "sheetData"}) {
			return "", fmt.Errorf("sheetData owner differs")
		}
		root, ok := data.Parent()
		if !ok || root != es[0] {
			return "", fmt.Errorf("worksheet owner differs")
		}
		var f, v losslessxml.Element
		formulaCount, valueCount := 0, 0
		for _, child := range es {
			owner, ok := child.Parent()
			if !ok || owner != cell {
				continue
			}
			switch child.Name() {
			case (xml.Name{Space: packaging.NSSpreadsheetML, Local: "f"}):
				f = child
				formulaCount++
			case (xml.Name{Space: packaging.NSSpreadsheetML, Local: "v"}):
				v = child
				valueCount++
			default:
				return "", fmt.Errorf("unclassified selected child %s!%s", sheet, address)
			}
		}
		if isInput {
			if formulaCount != 0 || valueCount != 1 || len(v.Attributes()) != 0 {
				return "", fmt.Errorf("input value topology differs")
			}
			text, leaf := v.Text()
			if !leaf || text != input[address] {
				return "", fmt.Errorf("input value differs %s/%q", address, text)
			}
			start, end := v.ContentRange()
			spans = append(spans, contract20LexicalRemoval{start, end, "[INPUT]"})
			continue
		}
		if formulaCount != 1 || len(f.Attributes()) != 0 {
			return "", fmt.Errorf("ordinary formula topology differs %s!%s", sheet, address)
		}
		body, leaf := f.Text()
		if !leaf || body != formulaWant.body {
			return "", fmt.Errorf("formula body differs %s!%s", sheet, address)
		}
		if saved {
			if valueCount != 0 {
				return "", fmt.Errorf("formula cache survived %s!%s", sheet, address)
			}
			continue
		}
		if valueCount != 1 || len(v.Attributes()) != 0 {
			return "", fmt.Errorf("source cache topology differs %s!%s", sheet, address)
		}
		cache, leaf := v.Text()
		if !leaf || cache != formulaWant.cache {
			return "", fmt.Errorf("source cache differs %s!%s", sheet, address)
		}
		start, end := v.SourceRange()
		spans = append(spans, contract20LexicalRemoval{start, end, ""})
	}
	if len(seen) != len(input)+len(formulas) {
		return "", fmt.Errorf("selected cells absent in %s: %v", sheet, seen)
	}
	// Sort in source order without synthesising a new worksheet.
	for i := 1; i < len(spans); i++ {
		for j := i; j > 0 && spans[j].start < spans[j-1].start; j-- {
			spans[j], spans[j-1] = spans[j-1], spans[j]
		}
	}
	return contract20MaskSpans(payload, spans)
}

func contract20WorkbookLexicalMask(original, changed []byte) error {
	parse := func(data []byte) (losslessxml.Element, error) {
		doc, err := losslessxml.Parse(data)
		if err != nil {
			return losslessxml.Element{}, err
		}
		es := doc.Elements()
		if len(es) == 0 || es[0].Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "workbook"}) {
			return losslessxml.Element{}, fmt.Errorf("workbook root differs")
		}
		var calc losslessxml.Element
		count := 0
		for _, e := range es {
			if e.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "calcPr"}) {
				continue
			}
			parent, ok := e.Parent()
			if !ok || parent != es[0] {
				return losslessxml.Element{}, fmt.Errorf("calcPr owner differs")
			}
			calc = e
			count++
		}
		if count != 1 {
			return losslessxml.Element{}, fmt.Errorf("calcPr count %d", count)
		}
		return calc, nil
	}
	oldCalc, err := parse(original)
	if err != nil {
		return err
	}
	newCalc, err := parse(changed)
	if err != nil {
		return err
	}
	const oldMarkup = `<calcPr calcId="124519"/>`
	if string(oldCalc.Raw()) != oldMarkup {
		return fmt.Errorf("original calcId/markup differs")
	}
	newRaw := string(newCalc.Raw())
	for _, flag := range []string{` calcMode="auto"`, ` fullCalcOnLoad="1"`, ` forceFullCalc="1"`} {
		if strings.Count(newRaw, flag) != 1 {
			return fmt.Errorf("calcPr flag %q not unique", flag)
		}
		newRaw = strings.Replace(newRaw, flag, "", 1)
	}
	if newRaw != oldMarkup {
		return fmt.Errorf("calcPr changed outside allowed flags: %s", newRaw)
	}
	start, end := newCalc.SourceRange()
	masked, err := contract20MaskSpans(changed, []contract20LexicalRemoval{{start, end, oldMarkup}})
	if err != nil {
		return err
	}
	if masked != string(original) {
		return fmt.Errorf("workbook intervening bytes/calcId differ")
	}
	return nil
}

// contract20CacheLexicalCustody validates both entire changed members against
// literal recipe bytes, excluding only the input value text, four complete old
// <v> cache elements, and the three new calcPr flags. Both sealed inputs have
// an existing calcPr, so this profile permits no insertion. Formula elements,
// cell attributes, calcId, whitespace and every intervening byte must survive.
func contract20CacheLexicalCustody(before, after map[string][]byte) error {
	for _, entry := range []struct{ part, sheet string }{{"xl/worksheets/sheet1.xml", "Model"}, {"xl/worksheets/sheet2.xml", "Summary"}} {
		old, ok := before[entry.part]
		if !ok {
			return fmt.Errorf("missing source member %s", entry.part)
		}
		newData, ok := after[entry.part]
		if !ok {
			return fmt.Errorf("missing saved member %s", entry.part)
		}
		a, e := contract20SheetLexicalMask(old, entry.sheet, false)
		if e != nil {
			return e
		}
		b, e := contract20SheetLexicalMask(newData, entry.sheet, true)
		if e != nil {
			return e
		}
		if a != b {
			return fmt.Errorf("unselected %s worksheet bytes differ", entry.sheet)
		}
	}
	return contract20WorkbookLexicalMask(before["xl/workbook.xml"], after["xl/workbook.xml"])
}

// Mutations affect valid XML inside one of the three allowed members. Each
// must be detected despite preserving the outer ZIP member allowance.
func TestContract20CacheLexicalMaskCollateralBatch(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	for _, recipe := range []string{"cross-caches", "opaque-caches"} {
		t.Run(recipe, func(t *testing.T) {
			state := &contract20ReadWorld{}
			if err := state.input(recipe); err != nil {
				t.Fatal(err)
			}
			s, err := spreadsheet.OpenEditing(state.source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			held, err := s.FindNumber("Model", "A1")
			if err != nil {
				t.Fatal(err)
			}
			if recipe == "opaque-caches" {
				_, err = s.SetNumberWithSealedOpaqueCaches(held, 10)
			} else {
				_, err = s.SetNumberWithAllCachesInvalidated(held, 10)
			}
			if err != nil {
				t.Fatal(err)
			}
			path := filepath.Join(t.TempDir(), "changed.xlsx")
			if _, err = s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			output, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			baseline, err := blankMemberPayloads(output)
			if err != nil {
				t.Fatal(err)
			}
			if err := contract20CacheLexicalCustody(state.parts, baseline); err != nil {
				t.Fatalf("unmodified result rejected: %v", err)
			}
			for _, mutation := range []struct{ name, part, old, new string }{
				{"input-value-attribute", "xl/worksheets/sheet1.xml", `<c r="A1"><v>10</v></c>`, `<c r="A1"><v custom="1">10</v></c>`},
				{"input-cell-attribute", "xl/worksheets/sheet1.xml", `<c r="A1"><v>10</v></c>`, `<c r="A1" s="0"><v>10</v></c>`},
				{"formula-body", "xl/worksheets/sheet1.xml", "<f>A1+A2</f>", "<f>A1-A2</f>"},
				{"formula-cache-reintroduced", "xl/worksheets/sheet1.xml", `<f>A1+A2</f></c>`, `<f>A1+A2</f><v>3</v></c>`},
				{"cell-attribute", "xl/worksheets/sheet1.xml", `<c r="B1">`, `<c r="B1" s="0">`},
				{"intervening-row", "xl/worksheets/sheet1.xml", `<row r="2">`, `<row r="2" custom="1">`},
				{"summary-formula", "xl/worksheets/sheet2.xml", "<f>40+2</f>", "<f>40+3</f>"},
				{"calc-id", "xl/workbook.xml", `calcId="124519"`, `calcId="124520"`},
				{"calc-extra-flag", "xl/workbook.xml", `calcId="124519"`, `calcId="124519" iterate="1"`},
			} {
				t.Run(mutation.name, func(t *testing.T) {
					candidate := make(map[string][]byte, len(baseline))
					for part, data := range baseline {
						candidate[part] = bytes.Clone(data)
					}
					original := string(candidate[mutation.part])
					if strings.Count(original, mutation.old) != 1 {
						t.Fatalf("mutation site %q not unique", mutation.old)
					}
					candidate[mutation.part] = []byte(strings.Replace(original, mutation.old, mutation.new, 1))
					if _, err := losslessxml.Parse(candidate[mutation.part]); err != nil {
						t.Fatalf("control not valid XML: %v", err)
					}
					if err := contract20CacheLexicalCustody(state.parts, candidate); err == nil {
						t.Error("collateral mutation passed exact lexical custody")
					}
				})
			}
		})
	}
}
