package formula

import (
	"reflect"
	"strings"
	"testing"
)

func TestStaticFormulaReferenceProperties(t *testing.T) {
	cells := []string{"A1", "$B2", "C$3", "$XFD$9"}
	prefixes := []string{"", "Main!", "'Input Data'!"}
	ops := []string{"+", "-", "*", "/", "&", ">="}
	for _, prefix := range prefixes {
		for _, a := range cells {
			for _, b := range cells {
				for _, op := range ops {
					source := prefix + a + op + b
					refs, err := Analyze(source)
					if err != nil || len(refs) != 2 {
						t.Fatalf("%s %+v %v", source, refs, err)
					}
					for _, r := range refs {
						raw := source[r.Start:r.End]
						again, err := Analyze(raw)
						if err != nil || len(again) != 1 || again[0].Sheet != r.Sheet || again[0].First != r.First || again[0].Last != r.Last {
							t.Fatalf("reference slice did not independently parse %q", raw)
						}
					}
					// Insertion below every referenced row is an exact lexical no-op.
					mapped, err := InsertReferences(source, "Main", Insertion{"Main", "row", 100, 1})
					if err != nil || mapped != source {
						t.Fatalf("unaffected remap changed %s", source)
					}
					// Wrapping a reference expression does not change its dependencies.
					wrapped, err := Analyze("SUM(" + source + ")")
					if err != nil {
						t.Fatal(err)
					}
					for i := range wrapped {
						wrapped[i].Start -= 4
						wrapped[i].End -= 4
					}
					if !reflect.DeepEqual(refs, wrapped) {
						t.Fatal("function wrapping changed dependencies")
					}
				}
			}
		}
	}
}
func FuzzStaticReferenceAnalysis(f *testing.F) {
	for _, source := range []string{`Input!A1*2`, `SUM($A$1:B3)+C4`, `IF(A1="+",B2,C3)`, `'O''Brien'!A1`, `INDIRECT(A1)`, `A1 "+" B1`} {
		f.Add(source)
	}
	f.Fuzz(func(t *testing.T, source string) {
		if len(source) > 65536 {
			return
		}
		refs, err := Analyze(source)
		if err != nil {
			if len(refs) != 0 {
				t.Fatal("partial references on refusal")
			}
			return
		}
		last := 0
		for _, r := range refs {
			if r.Start < last || r.End <= r.Start || r.End > len(source) {
				t.Fatal("invalid reference range")
			}
			last = r.End
			if r.First.Row < 1 || r.First.Column < 1 || r.Last.Row > 1048576 || r.Last.Column > 16384 {
				t.Fatal("unbounded coordinates")
			}
		}
		output, err := InsertReferences(source, "Main", Insertion{"Other", "row", 1, 1})
		if err != nil {
			if output != "" {
				t.Fatal("partial remap")
			}
			return
		}
		if _, err = Analyze(output); err != nil {
			t.Fatal("remap produced unsupported expression")
		}
		if !strings.Contains(strings.ToLower(source), "other!") && output != source {
			t.Fatal("foreign sheet remap changed formula")
		}
	})
}
