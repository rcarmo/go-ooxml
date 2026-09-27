package formula

import "testing"

func TestStaticFormulaAnalysisBatch(t *testing.T) {
	for _, tc := range []struct {
		source string
		count  int
	}{{`IF(A1="B2",'O''Brien'!$C$4,SUM(D1:E2))`, 3}, {`"A1"&"Sheet1!B2"`, 0}, {`LOG10(A1)+1E10`, 1}, {`=-(A1+B2)^2%`, 2}, {`TRUE`, 0}, {`'α sheet'!A1`, 1}} {
		t.Run(tc.source, func(t *testing.T) {
			refs, err := Analyze(tc.source)
			if err != nil || len(refs) != tc.count {
				t.Fatalf("%+v %v", refs, err)
			}
			for _, r := range refs {
				if r.Start < 0 || r.End <= r.Start || r.End > len(tc.source) {
					t.Fatal("bad lexical offsets")
				}
			}
		})
	}
	for _, source := range []string{`NOW()`, `MyFunction(A1)`, `Name+1`, `XFE1`, `A1048577`, `A0`, `A1:B`, `$A$`, `SUM(A1,)`, `SUM(,A1)`, `1..2`, `"unclosed`, `'unclosed!A1`, `A1::B2`, `A1;B2`, `A1!B2!C3`, `SUM()`, `=`, `@A1`, `A1+SUM(2`} {
		t.Run(source, func(t *testing.T) {
			refs, err := Analyze(source)
			if err == nil || len(refs) != 0 {
				t.Fatalf("accepted %q %+v", source, refs)
			}
		})
	}
	t.Run("quoted sheet absolute flags offsets", func(t *testing.T) {
		source := `'O''Brien'!$B2:C$4`
		refs, err := Analyze(source)
		if err != nil || len(refs) != 1 {
			t.Fatal(refs, err)
		}
		r := refs[0]
		if r.Sheet != "O'Brien" || r.First.Column != 2 || r.First.Row != 2 || !r.First.AbsoluteColumn || r.First.AbsoluteRow || r.Last.Column != 3 || r.Last.Row != 4 || !r.Last.AbsoluteRow || r.Start != 0 || r.End != len(source) {
			t.Fatalf("%+v", r)
		}
	})
}
