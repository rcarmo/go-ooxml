package spreadsheet

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestContract20ReadOrdinaryFormulaCacheBatch(t *testing.T) {
	for _, recipe := range []string{"cross-caches", "opaque-caches"} {
		t.Run(recipe, func(t *testing.T) {
			input := contract20ReadInput(t, recipe)
			prior := contract20InputMembers(t, input)
			original := bytes.Clone(input)
			s, err := OpenEditing(input, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			for _, item := range []struct {
				sheet, address string
				prior          float64
			}{
				{"Model", "B1", 3}, {"Model", "B2", 6}, {"Summary", "A1", 9}, {"Summary", "B1", 42},
			} {
				value, err := s.ReadFormulaCache(item.sheet, item.address)
				if err != nil || value == nil || *value != item.prior {
					t.Errorf("prior %s!%s = %v/%v", item.sheet, item.address, value, err)
				}
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
			for _, item := range []struct{ sheet, address string }{{"Model", "B1"}, {"Model", "B2"}, {"Summary", "A1"}, {"Summary", "B1"}} {
				cache, e := s.ReadFormulaCache(item.sheet, item.address)
				if e != nil || cache != nil {
					t.Errorf("cleared %s!%s = %v/%v", item.sheet, item.address, cache, e)
				}
			}
			path := filepath.Join(t.TempDir(), "edited.xlsx")
			if _, err = s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			reopened, err := OpenEditing(saved, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			for _, item := range []struct{ sheet, address string }{{"Model", "B1"}, {"Model", "B2"}, {"Summary", "A1"}, {"Summary", "B1"}} {
				cache, e := reopened.ReadFormulaCache(item.sheet, item.address)
				if e != nil || cache != nil {
					t.Errorf("reopened %s!%s = %v/%v", item.sheet, item.address, cache, e)
				}
			}
			if !bytes.Equal(input, original) || !reflect.DeepEqual(prior, contract20InputMembers(t, input)) {
				t.Error("read/edit changed caller archive")
			}
		})
	}
}
func TestContract20ReadFormulaCacheRefusalBatch(t *testing.T) {
	for _, profile := range []struct{ recipe, sheet, address, kind string }{
		{"attributed-formulas", "Calc", "B2", "xlsx-cache-topology-unsupported"},
		{"attributed-formulas", "Calc", "D2", "xlsx-cache-topology-unsupported"},
		{"cross-caches", "Model", "A1", "unsupported_structure"},
		{"cross-caches", "Model", "ZZ9", "missing_target"},
	} {
		t.Run(profile.recipe+"/"+profile.address, func(t *testing.T) {
			input := contract20ReadInput(t, profile.recipe)
			prior := contract20InputMembers(t, input)
			s, err := OpenEditing(input, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			value, err := s.ReadFormulaCache(profile.sheet, profile.address)
			var refusal *packaging.Refusal
			if value != nil || !errors.As(err, &refusal) || refusal.Kind != profile.kind {
				t.Errorf("read %v/%v want %s", value, err, profile.kind)
			}
			if !reflect.DeepEqual(prior, contract20InputMembers(t, input)) {
				t.Error("refusal changed input members")
			}
		})
	}
	// A foreign/attributed value is not silently exposed as numeric cache.
	input := contract20ReadInput(t, "cross-caches")
	modified := contract20AlterMember(t, input, "xl/worksheets/sheet1.xml", func(old string) string {
		return strings.Replace(old, "<f>A1+A2</f>", `<f t="array" ref="B1:B2">A1+A2</f>`, 1)
	})
	s, err := OpenEditing(modified, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	if v, e := s.ReadFormulaCache("Model", "B1"); v != nil || e == nil {
		t.Errorf("attributed read exposed cache %v/%v", v, e)
	}
}
