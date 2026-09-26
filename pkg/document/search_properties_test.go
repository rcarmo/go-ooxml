package document

import "testing"

// Fixed seed corpus runs in normal/race batches; exploratory fuzz jobs remain a
// separate integration gate and are not represented by a single-test invocation.
func FuzzMappedSearchIntervals(f *testing.F) {
	for _, s := range []string{"Straße – coût", "sß", "A  B", "İΣς", "co\u00adoperate", "😀é"} {
		f.Add(s, "s")
	}
	f.Fuzz(func(t *testing.T, text, needle string) {
		if len(text) > 32768 || len(needle) > 1024 {
			return
		}
		raw := []rune(text)
		folded, mapping := matchSpace(raw, true)
		n, _ := matchSpace([]rune(needle), true)
		for _, span := range matchIntervals(folded, n, mapping) {
			if span[0] < 0 || span[1] <= span[0] || span[1] > len(raw) {
				t.Fatalf("invalid source bounds %+v", span)
			}
			selected, _ := matchSpace(raw[span[0]:span[1]], true)
			if len(selected) < len(n) {
				t.Fatalf("selected fewer folded characters than match")
			}
		}
	})
}
