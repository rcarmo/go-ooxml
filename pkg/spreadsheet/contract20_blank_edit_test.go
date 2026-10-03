package spreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// A deliberately unselected native control changes only literal fixture input
// members. It cannot supply the implementation with a generated output oracle.
func contract20AlternateBlankInput(t *testing.T) []byte {
	t.Helper()
	members := contract20InputMembers(t, contract20ReadInput(t, "prefixed-cells"))
	workbook := string(members["xl/workbook.xml"])
	if !strings.Contains(workbook, `name="Prefixed"`) {
		t.Fatal("pinned workbook shape drift")
	}
	members["xl/workbook.xml"] = []byte(strings.Replace(workbook, `name="Prefixed"`, `name="OtherSheet"`, 1))
	sheet := string(members["xl/worksheets/sheet1.xml"])
	if !strings.Contains(sheet, `<x:row r="1"><x:c r="A1" s="1"/><x:c r="B1" s="2"/><x:c r="C1"><x:f>1+1</x:f><x:v>2</x:v></x:c></x:row>`) {
		t.Fatal("pinned sheet shape drift")
	}
	sheet = strings.Replace(sheet, `<x:row r="1"><x:c r="A1" s="1"/><x:c r="B1" s="2"/><x:c r="C1"><x:f>1+1</x:f><x:v>2</x:v></x:c></x:row>`, `<x:row r="4"><x:c r="D4" s="1"/><x:c r="E4" s="2"/><x:c r="F4"><x:f>D4*3</x:f><x:v>18</x:v></x:c></x:row>`, 1)
	members["xl/worksheets/sheet1.xml"] = []byte(sheet)
	return contract20ZipMembers(t, members)
}

func contract20ZipMembers(t *testing.T, members map[string][]byte) []byte {
	t.Helper()
	var out bytes.Buffer
	writer := zip.NewWriter(&out)
	parts := make([]string, 0, len(members))
	for part := range members {
		parts = append(parts, part)
	}
	sort.Strings(parts)
	for _, part := range parts {
		w, err := writer.Create(part)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(members[part]); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func TestContract20BlankAlternateNativeControl(t *testing.T) {
	input := contract20AlternateBlankInput(t)
	before := contract20InputMembers(t, input)
	edits := []ContractBlankEdit{{Address: "D4", Value: `Another & <value>`}, {Address: "E4", Value: float64(23.5)}}
	s, err := OpenEditing(input, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	if err := s.SetContractBlankValues("OtherSheet", edits); err != nil {
		t.Fatalf("alternate operands refused: %v", err)
	}
	path := filepath.Join(t.TempDir(), "alternate.xlsx")
	if _, err := s.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	output, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	after := contract20InputMembers(t, output)
	for part, prior := range before {
		if part != "xl/workbook.xml" && part != "xl/worksheets/sheet1.xml" && !bytes.Equal(prior, after[part]) {
			t.Errorf("unselected member changed: %s", part)
		}
	}
	if !reflect.DeepEqual(before, contract20InputMembers(t, input)) {
		t.Error("input members mutated")
	}
	root := string(after["xl/worksheets/sheet1.xml"])
	for _, want := range []string{`<x:c r="D4" s="1" t="inlineStr"><x:is><x:t>Another &amp; &lt;value&gt;</x:t></x:is></x:c>`, `<x:c r="E4" s="2"><x:v>23.5</x:v></x:c>`, `<x:c r="F4"><x:f>D4*3</x:f></x:c>`} {
		if !strings.Contains(root, want) {
			t.Errorf("missing retained edited fragment %s: %s", want, root)
		}
	}
	if strings.Contains(root, `<x:v>18</x:v>`) {
		t.Error("alternate formula cache retained")
	}
	wbXML := string(after["xl/workbook.xml"])
	for _, want := range []string{`<x:calcPr`, `calcMode="auto"`, `fullCalcOnLoad="1"`, `forceFullCalc="1"`} {
		if !strings.Contains(wbXML, want) {
			t.Errorf("missing recalculation metadata %s", want)
		}
	}
	wb, err := OpenReader(bytes.NewReader(output), int64(len(output)))
	if err != nil {
		t.Fatal(err)
	}
	defer wb.Close()
	opened, err := wb.Sheet("OtherSheet")
	if err != nil {
		t.Fatal(err)
	}
	if got := opened.Cell("D4").String(); got != `Another & <value>` {
		t.Errorf("D4 reopened %q", got)
	}
	if got, e := opened.Cell("E4").Float64(); e != nil || got != 23.5 {
		t.Errorf("E4 reopened %v/%v", got, e)
	}
	// Attributed/shared formula topology must refuse before mutating the held session.
	broken := contract20InputMembers(t, input)
	broken["xl/worksheets/sheet1.xml"] = []byte(strings.Replace(string(broken["xl/worksheets/sheet1.xml"]), `<x:f>D4*3</x:f>`, `<x:f t="shared" si="0">D4*3</x:f>`, 1))
	refused := contract20ZipMembers(t, broken)
	guard, err := OpenEditing(refused, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	err = guard.SetContractBlankValues("OtherSheet", edits)
	var rejection *packaging.Refusal
	if !errors.As(err, &rejection) || rejection.Kind != "xlsx-cache-topology-unsupported" {
		t.Errorf("attributed formula refusal %v", err)
	}
	path = filepath.Join(t.TempDir(), "refused.xlsx")
	if _, err := guard.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	saved, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(broken, contract20InputMembers(t, saved)) || !reflect.DeepEqual(broken, contract20InputMembers(t, refused)) {
		t.Error("refusal altered input or saved members")
	}
}

func TestContract20BlankEditBatch(t *testing.T) {
	for _, tc := range []struct {
		name, sheet string
		edits       []ContractBlankEdit
		changed     []string
	}{
		{"styled-blank", "StyledBlank", []ContractBlankEdit{{Address: "A1", Value: "filled"}}, []string{"xl/worksheets/sheet1.xml"}},
		{"prefixed-cells", "Prefixed", []ContractBlankEdit{{Address: "A1", Value: "Alpha"}, {Address: "B1", Value: float64(7)}}, []string{"xl/workbook.xml", "xl/worksheets/sheet1.xml"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := contract20ReadInput(t, tc.name)
			original := append([]byte(nil), source...)
			before := contract20InputMembers(t, source)
			session, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if err := session.SetContractBlankValues(tc.sheet, tc.edits); err != nil {
				t.Fatalf("selected blank write refused: %v", err)
			}
			dest := filepath.Join(t.TempDir(), "edited.xlsx")
			if _, err := session.SaveAs(dest); err != nil {
				t.Fatal(err)
			}
			output, err := os.ReadFile(dest)
			if err != nil {
				t.Fatal(err)
			}
			after := contract20InputMembers(t, output)
			if len(after) != len(before) {
				t.Errorf("member count %d want %d", len(after), len(before))
			}
			var changed []string
			for part, old := range before {
				now, ok := after[part]
				if !ok {
					t.Errorf("member disappeared %s", part)
				} else if !bytes.Equal(old, now) {
					changed = append(changed, part)
				}
			}
			sort.Strings(changed)
			if !reflect.DeepEqual(changed, tc.changed) {
				t.Errorf("changed members %v, want %v", changed, tc.changed)
			}
			wb, err := OpenReader(bytes.NewReader(output), int64(len(output)))
			if err != nil {
				t.Fatal(err)
			}
			defer wb.Close()
			sheet, err := wb.Sheet(tc.sheet)
			if err != nil {
				t.Fatal(err)
			}
			if got := sheet.Cell("A1").String(); got != tc.edits[0].Value {
				t.Errorf("reopened A1 %q, want %v", got, tc.edits[0].Value)
			}
			if tc.name == "styled-blank" {
				if got, e := sheet.Cell("B1").Float64(); e != nil || got != 5 {
					t.Errorf("B1 neighbour %v/%v", got, e)
				}
			} else {
				if got, e := sheet.Cell("B1").Float64(); e != nil || got != 7 {
					t.Errorf("B1 numeric %v/%v", got, e)
				}
				root := string(after["xl/worksheets/sheet1.xml"])
				if !strings.Contains(root, `<x:is><x:t>Alpha</x:t></x:is>`) || !strings.Contains(root, `<x:v>7</x:v>`) || strings.Contains(root, `<x:v>2</x:v>`) {
					t.Errorf("prefixed values/cache not qualified/cleared: %s", root)
				}
				var workbook struct {
					Calc []struct {
						Mode, Full, Force string `xml:",attr"`
					} `xml:"calcPr"`
				}
				if err := xml.Unmarshal(after["xl/workbook.xml"], &workbook); err != nil {
					t.Fatal(err)
				}
				text := string(after["xl/workbook.xml"])
				if strings.Count(text, "<x:calcPr") != 1 || !strings.Contains(text, `calcMode="auto"`) || !strings.Contains(text, `fullCalcOnLoad="1"`) || !strings.Contains(text, `forceFullCalc="1"`) {
					t.Errorf("qualified calcPr missing: %s", text)
				}
				if !strings.Contains(root, `<x:f>1+1</x:f>`) {
					t.Error("C1 formula changed")
				}
			}
			if !bytes.Equal(original, source) || !reflect.DeepEqual(before, contract20InputMembers(t, source)) {
				t.Error("source/caller bytes changed")
			}
			if !bytes.Equal(before["xl/styles.xml"], after["xl/styles.xml"]) {
				t.Error("style registry changed")
			}
			oldSheet, newSheet := string(before["xl/worksheets/sheet1.xml"]), string(after["xl/worksheets/sheet1.xml"])
			// Independent lexical custody outside entire selected cells and the sole C1 cache.
			blank := func(s string, ref string) string {
				var doc struct {
					Cells []struct {
						R string `xml:"r,attr"`
					} `xml:"sheetData>row>c"`
				}
				if err := xml.Unmarshal([]byte(s), &doc); err != nil {
					t.Fatal(err)
				}
				found := 0
				for _, cell := range doc.Cells {
					if cell.R == ref {
						found++
					}
				}
				if found != 1 {
					t.Fatalf("selected cell %s count %d", ref, found)
				}
				start := strings.Index(s, ` r="`+ref+`"`)
				if start < 0 {
					t.Fatalf("missing cell %s", ref)
				}
				left := strings.LastIndex(s[:start], "<")
				tagEnd := strings.Index(s[start:], ">")
				if tagEnd < 0 {
					t.Fatalf("unterminated start tag %s", ref)
				}
				end := start + tagEnd + 1
				if s[end-2] == '/' {
					return s[:left] + "[CELL]" + s[end:]
				}
				name := "c"
				if strings.HasPrefix(s[left:], "<x:c ") {
					name = "x:c"
				}
				closeTag := "</" + name + ">"
				closeAt := strings.Index(s[end:], closeTag)
				if closeAt < 0 {
					t.Fatalf("unterminated cell %s", ref)
				}
				return s[:left] + "[CELL]" + s[end+closeAt+len(closeTag):]
			}
			for _, e := range tc.edits {
				oldSheet = blank(oldSheet, e.Address)
				newSheet = blank(newSheet, e.Address)
			}
			if tc.name == "prefixed-cells" {
				oldSheet = strings.Replace(oldSheet, `<x:v>2</x:v>`, `[CACHE]`, 1)
				newSheet = strings.Replace(newSheet, `<x:f>1+1</x:f>`, `<x:f>1+1</x:f>[CACHE]`, 1)
			}
			if oldSheet != newSheet {
				t.Errorf("unselected worksheet lexical bytes differ\nold %s\nnew %s", oldSheet, newSheet)
			}
		})
	}
}
