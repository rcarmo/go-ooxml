package formula

import "testing"

func TestReferenceInsertionBatch(t *testing.T) {
	for _, tc := range []struct {
		source, want string
		change       Insertion
	}{{`IF(A1="A2",A2,Other!A2)`, `IF(A1="A2",A3,Other!A2)`, Insertion{"Main", "row", 2, 1}}, {`'O''Brien'!$b$2 + Main!c1`, `'O''Brien'!$b$2 + Main!D1`, Insertion{"Main", "column", 2, 1}}, {`SUM(A5:A1)`, `SUM(A7:A1)`, Insertion{"Main", "row", 3, 2}}, {`a1 + Other!b2`, `a1 + Other!b2`, Insertion{"Main", "row", 3, 2}}, {`Main!A1:A3`, `Main!A1:A4`, Insertion{"main", "row", 2, 1}}} {
		t.Run(tc.source, func(t *testing.T) {
			out, err := InsertReferences(tc.source, "Main", tc.change)
			if err != nil || out != tc.want {
				t.Fatalf("%q %v", out, err)
			}
		})
	}
	for _, change := range []Insertion{{"Main", "bad", 1, 1}, {"Main", "row", 0, 1}, {"Main", "column", 1, 0}, {"", "row", 1, 1}, {"Main", "row", 1048577, 1}, {"Main", "column", 1, 16385}} {
		if out, err := InsertReferences(`A1+B1`, "Main", change); err == nil || out != "" {
			t.Fatal("invalid remap accepted")
		}
	}
	for _, source := range []string{`XFD1`, `A1048576`, `INDIRECT(A1)`} {
		axis := "row"
		if source == "XFD1" {
			axis = "column"
		}
		if out, err := InsertReferences(source, "Main", Insertion{"Main", axis, 1, 1}); err == nil || out != "" {
			t.Fatal("unproved remap accepted")
		}
	}
}
