package spreadsheet

import (
	"bytes"
	"encoding/xml"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestContract20XLSXCacheProfiles(t *testing.T) {
	for _, recipe := range []string{"cross-caches", "input-array", "input-dataTable", "opaque-caches"} {
		t.Run(recipe, func(t *testing.T) {
			source := contract20ReadInput(t, recipe)
			original := append([]byte(nil), source...)
			originalParts := contract20InputMembers(t, source)
			session, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			target, err := session.FindNumber("Model", "A1")
			if err != nil {
				t.Fatal(err)
			}
			if recipe == "input-array" || recipe == "input-dataTable" {
				result, err := session.SetNumberWithAllCachesInvalidated(target, 10)
				var refusal *packaging.Refusal
				if !errors.As(err, &refusal) || refusal.Kind != "xlsx-cache-topology-unsupported" {
					t.Errorf("array/dataTable preflight = %#v, effect %#v", err, result)
				}
				if !bytes.Equal(source, original) || !reflect.DeepEqual(originalParts, contract20InputMembers(t, source)) {
					t.Error("refusal changed original/caller bytes")
				}
				refusedPath := filepath.Join(t.TempDir(), "refused.xlsx")
				if _, err := session.SaveAs(refusedPath); err != nil {
					t.Fatal(err)
				}
				refused, err := os.ReadFile(refusedPath)
				if err != nil {
					t.Fatal(err)
				}
				if !reflect.DeepEqual(originalParts, contract20InputMembers(t, refused)) {
					t.Error("refused save changed member payloads")
				}
				for _, payload := range []byte{1, 2} {
					if !strings.Contains(string(originalParts["xl/worksheets/sheet2.xml"]), "<v>"+strconv.Itoa(int(payload))+"</v>") {
						t.Error("missing original array/dataTable cache", payload)
					}
				}
				return
			}
			edit := session.SetNumberWithAllCachesInvalidated
			if recipe == "opaque-caches" {
				edit = session.SetNumberWithSealedOpaqueCaches
			}
			// Independent no-op session: no formula cache or calculation flag is rewritten.
			if effect, err := edit(target, 1); err != nil || effect.ValueChanged || len(effect.Invalidated) != 0 {
				t.Errorf("no-op = %#v / %v", effect, err)
			}
			noOpPath := filepath.Join(t.TempDir(), "noop.xlsx")
			if _, err := session.SaveAs(noOpPath); err != nil {
				t.Fatal(err)
			}
			noOp, err := os.ReadFile(noOpPath)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(originalParts, contract20InputMembers(t, noOp)) {
				t.Error("no-op member payload or flag drift")
			}
			result, err := edit(target, 10)
			if err != nil {
				t.Fatal(err)
			}
			wantInvalidated := []string{"Model!B1", "Model!B2", "Summary!A1", "Summary!B1"}
			if result.State != "recalculation-required" || !result.ValueChanged || !reflect.DeepEqual(result.Invalidated, wantInvalidated) {
				t.Errorf("invalidate-all effect = %#v; want %v", result, wantInvalidated)
			}
			path := filepath.Join(t.TempDir(), "changed.xlsx")
			if _, err := session.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			data, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			parts := contract20InputMembers(t, data)
			if len(parts) != len(originalParts) {
				t.Errorf("member count %d, want %d", len(parts), len(originalParts))
			}
			var changed []string
			for name, before := range originalParts {
				after, ok := parts[name]
				if !ok {
					t.Errorf("lost member %s", name)
				} else if !bytes.Equal(before, after) {
					changed = append(changed, name)
				}
			}
			sort.Strings(changed)
			wantChanged := []string{"xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"}
			if !reflect.DeepEqual(changed, wantChanged) {
				t.Errorf("changed members %v, want %v", changed, wantChanged)
			}
			workbook := string(parts["xl/workbook.xml"])
			for _, literal := range []string{`calcMode="auto"`, `fullCalcOnLoad="1"`, `forceFullCalc="1"`} {
				if !strings.Contains(workbook, literal) {
					t.Errorf("missing calc flag %s", literal)
				}
			}
			for _, sheet := range []string{"xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"} {
				if strings.Contains(string(parts[sheet]), "<v>42</v>") {
					t.Error("unrelated cache 42 retained", sheet)
				}
			}
			if !bytes.Equal(source, original) || !reflect.DeepEqual(originalParts, contract20InputMembers(t, source)) {
				t.Error("edit changed input/caller bytes")
			}
			expected := map[string]map[string]string{
				"xl/worksheets/sheet1.xml": {"B1": "A1+A2", "B2": "B1*2"},
				"xl/worksheets/sheet2.xml": {"A1": "Model!B1*3", "B1": "40+2"},
			}
			for part, formulas := range expected {
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
					t.Fatal(err)
				}
				seen := map[string]bool{}
				for _, row := range sheet.Rows {
					for _, cell := range row.Cells {
						if want, ok := formulas[cell.Address]; ok {
							seen[cell.Address] = true
							if cell.Formula != want || cell.Value != nil {
								t.Errorf("%s!%s formula/cache = %q/%v, want %q/nil", part, cell.Address, cell.Formula, cell.Value, want)
							}
						}
					}
				}
				if len(seen) != len(formulas) {
					t.Errorf("%s formula cells seen %v", part, seen)
				}
			}
			reopened, err := OpenReader(bytes.NewReader(data), int64(len(data)))
			if err != nil {
				t.Fatal(err)
			}
			defer reopened.Close()
			model, err := reopened.Sheet("Model")
			if err != nil {
				t.Fatal(err)
			}
			if got, err := model.Cell("A1").Float64(); err != nil || got != 10 {
				t.Errorf("reopened Model!A1 %v/%v", got, err)
			}
		})
	}
}
