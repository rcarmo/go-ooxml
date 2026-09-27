package formula

import "testing"

func TestStaticRangeParserBatch(t *testing.T) {
	for _, tc := range []struct {
		text, sheet, axis string
		row, col          int
	}{{"$B:$B", "", "col", 0, 2}, {"$3:$1", "", "row", 3, 0}, {"'O''Brien'!$B:$B", "O'Brien", "col", 0, 2}, {"A3:B1", "", "cell", 3, 1}, {"Sheet!$A$1", "Sheet", "cell", 1, 1}, {"'3'!1:1", "3", "row", 1, 0}} {
		t.Run(tc.text, func(t *testing.T) {
			r, err := ParseRange(tc.text)
			if err != nil || r.Sheet != tc.sheet || r.First.Row != tc.row || r.First.Column != tc.col || r.WholeColumns != (tc.axis == "col") || r.WholeRows != (tc.axis == "row") {
				t.Fatalf("%+v %v", r, err)
			}
		})
	}
	for _, text := range []string{"IN", "1", "$A", "$1", "A1:A", "A:B1", "A:A+1", "SUM(A1)", "[Book]Sheet!A1", "Sheet1:Sheet2!A1", "'x'!A1!B1", "A1,B2", "A1 B2", "A1#", "XFE:XFE", "0:1", "1048577:1048577", "1E2:2", "\"A1\"", "", "=A1"} {
		t.Run(text, func(t *testing.T) {
			if _, err := ParseRange(text); err == nil {
				t.Fatal("unproved range accepted")
			}
		})
	}
}
