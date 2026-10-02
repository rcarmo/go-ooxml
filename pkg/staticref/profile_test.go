package staticref

import (
	"strings"
	"testing"
)

// Public profile boundary must exist; legacy internal grammar is not an external API.
func TestUniformStaticProfileRedBatch(t *testing.T) {
	refs, err := Analyze(`SUM(A1)+B2`)
	if err != nil || len(refs) != 2 || refs[0].Start != 4 || refs[0].End != 6 {
		t.Fatalf("reference profile: %+v %v", refs, err)
	}
	if got, err := Analyze(`IFERROR(A1,0)`); err == nil || got != nil {
		t.Fatalf("nonintersection function accepted: %+v %v", got, err)
	}
	if got, err := Analyze(`SUM(A1,)`); err == nil || got != nil {
		t.Fatalf("blank argument accepted: %+v %v", got, err)
	}
	if got, err := ParseRange(`$3:$1`); err != nil || got.First.Row != 3 || got.Last.Row != 1 {
		t.Fatalf("direct range: %+v %v", got, err)
	}
	if got, err := InsertReferences(`A2`, "Main", Insertion{Sheet: "Main", Axis: "row", At: 2, Count: 1}); err != nil || got != "A3" {
		t.Fatalf("remap: %q %v", got, err)
	}
}

func TestUniformStaticBoundaryBatch(t *testing.T) {
	for _, source := range []string{"A١", "$A1!B2", "A$1!B2", "Ⅳ!A1", "A1\t+B2", "A1\u00a0+B2", "A1 : B2", "'Main' !A1", "\"x\x01\""} {
		refs, err := Analyze(source)
		if refs != nil || !isCategory(err, "unsupported-static-reference") {
			t.Errorf("analysis %q: %+v %v", source, refs, err)
		}
	}
	for _, source := range []string{" A1", "A1 ", "A1 : B2", "A١", "$A1!B2", "A$1!B2", "Ⅳ!A1", "A1\t"} {
		r, err := ParseRange(source)
		if r != (Range{}) || !isCategory(err, "unsupported-direct-range") {
			t.Errorf("range %q: %+v %v", source, r, err)
		}
	}
	if refs, err := Analyze("AⅣ!A1"); err != nil || len(refs) != 1 || refs[0].Sheet != "AⅣ" || refs[0].Start != 0 || refs[0].End != 7 || refs[0].First != (Cell{Row: 1, Column: 1}) {
		t.Errorf("Unicode number sheet identifier %+v %v", refs, err)
	}
	if r, err := ParseRange("AⅣ!A1"); err != nil || r.Sheet != "AⅣ" || r.First != (Cell{Row: 1, Column: 1}) || r.Last != r.First {
		t.Errorf("Unicode direct sheet %+v %v", r, err)
	}
	if refs, err := Analyze("'AⅣ'!A1"); err != nil || len(refs) != 1 || refs[0].Sheet != "AⅣ" {
		t.Errorf("quoted Unicode sheet %+v %v", refs, err)
	}
	if r, err := ParseRange("'AⅣ'!A1"); err != nil || r.Sheet != "AⅣ" {
		t.Errorf("quoted direct sheet %+v %v", r, err)
	}
	if refs, err := Analyze("'Ⅳ'!A1"); err != nil || len(refs) != 1 || refs[0].Sheet != "Ⅳ" {
		t.Errorf("quoted numeral sheet %+v %v", refs, err)
	}
	if r, err := ParseRange("'A$1'!B2"); err != nil || r.Sheet != "A$1" {
		t.Errorf("quoted dollar sheet %+v %v", r, err)
	}
	// A quoted sheet is one lexer token, not one token per Unicode scalar.
	unicodeSheet := "'" + strings.Repeat("α", 50001) + "'!A1"
	if refs, err := Analyze(unicodeSheet); err != nil || len(refs) != 1 || refs[0].Sheet != strings.Repeat("α", 50001) {
		t.Errorf("Unicode token budget: refs %d err %v", len(refs), err)
	}
	if _, err := ParseRange(strings.Repeat("A", 1048577)); !isCategory(err, "unsupported-direct-range") {
		t.Errorf("direct limit: %v", err)
	}
	for _, depth := range []int{128, 129} {
		refs, err := Analyze(strings.Repeat("(", depth) + "A1" + strings.Repeat(")", depth))
		if depth == 128 && (err != nil || len(refs) != 1) || depth == 129 && (refs != nil || !isCategory(err, "static-reference-limit")) {
			t.Errorf("depth %d: %+v %v", depth, refs, err)
		}
	}
	change := Insertion{Sheet: "main", Axis: "row", At: 2, Count: 1}
	if out, err := InsertReferences("a1:b2", "Main", change); err != nil || out != "A1:B3" {
		t.Errorf("range render: %q %v", out, err)
	}
	if out, err := InsertReferences("a1:b2", "Main", Insertion{Sheet: "Main", Axis: "row", At: 100, Count: 1}); err != nil || out != "a1:b2" {
		t.Errorf("unchanged range spelling: %q %v", out, err)
	}
	for _, sheet := range []string{"Bad/Sheet", strings.Repeat("a", 256), "x\x7f"} {
		if out, err := InsertReferences("A1", "Main", Insertion{Sheet: sheet, Axis: "row", At: 1, Count: 1}); out != "" || !isCategory(err, "invalid-reference-insertion") {
			t.Errorf("sheet %q: %q %v", sheet, out, err)
		}
	}
	if out, err := InsertReferences(strings.Repeat("A1+", 260000)+"A1", "Main", change); out != "" || err == nil {
		t.Errorf("oversize output/source: %v", err)
	}
}
func isCategory(err error, category string) bool {
	e, ok := err.(*Refusal)
	return ok && e.Category == category
}
