package document

import "testing"

func TestNormalisedOverlappingCandidateBatch(t *testing.T) {
	for _, tc := range []struct{ raw, needle, want string }{{"sß", "ss", "ß"}, {"fﬃ", "ffi", "ﬃ"}, {"iİ", "i̇", "İ"}} {
		t.Run(tc.raw, func(t *testing.T) {
			raw := []rune(tc.raw)
			v, m := matchSpace(raw, true)
			n, _ := matchSpace([]rune(tc.needle), true)
			got := matchIntervals(v, n, m)
			if len(got) != 1 || string(raw[got[0][0]:got[0][1]]) != tc.want {
				t.Fatalf("%v", got)
			}
		})
	}
}
