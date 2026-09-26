package document

import "testing"

func TestNormalisedMappingBatch(t *testing.T) {
	for _, tc := range []struct{ input, want string }{{"Straße", "strasse"}, {"İΣςſKﬃ", "i̇σσskffi"}, {"“Cost” – ‘X’\u00a0\tY", "\"cost\" - 'x' y"}, {"co\u00adoperate", "cooperate"}, {"\x1cA\u2007B", " a b"}} {
		t.Run(tc.input, func(t *testing.T) {
			out, m := matchSpace([]rune(tc.input), true)
			if string(out) != tc.want || len(out) != len(m) {
				t.Fatalf("%q %v", out, m)
			}
		})
	}
	for _, tc := range []struct {
		raw, needle string
		count       int
		selected    string
	}{{"ß", "s", 0, ""}, {"ß", "ss", 1, "ß"}, {"İ", "i", 0, ""}, {"A  B", "a b", 1, "A  B"}, {"co\u00adop", "coop", 1, "co\u00adop"}, {"aaa", "aa", 1, "aa"}} {
		t.Run(tc.raw+tc.needle, func(t *testing.T) {
			raw := []rune(tc.raw)
			v, m := matchSpace(raw, true)
			n, _ := matchSpace([]rune(tc.needle), true)
			matches := matchIntervals(v, n, m)
			if len(matches) != tc.count {
				t.Fatal(matches)
			}
			if len(matches) > 0 && string(raw[matches[0][0]:matches[0][1]]) != tc.selected {
				t.Fatal("mapping lost source text")
			}
		})
	}
	if len(fullCaseFold) < 1500 {
		t.Fatalf("incomplete Unicode table: %d", len(fullCaseFold))
	}
}
