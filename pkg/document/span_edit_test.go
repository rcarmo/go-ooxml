package document

import "testing"

func TestAffixLocalizationBatch(t *testing.T) {
	for _, tc := range []struct {
		old, new     string
		lo, hi, next int
		ambiguous    bool
	}{{"Alpha", "Omega", 0, 4, 4, false}, {"Alpha", "Beta", 0, 4, 3, false}, {"TermTerm", "Term", 0, 0, 0, true}, {"abc", "abXc", 2, 2, 3, false}, {"aaa", "aa", 0, 0, 0, true}, {"😀café", "tea", 0, 5, 3, false}, {"pre OLD post", "pre NEW post", 4, 7, 7, false}, {"abc", "", 0, 3, 0, false}} {
		t.Run(tc.old+"_"+tc.new, func(t *testing.T) {
			lo, hi, n, err := changedInterval([]rune(tc.old), []rune(tc.new))
			if (err != nil) != tc.ambiguous {
				t.Fatalf("ambiguity %v", err)
			}
			if err == nil && (lo != tc.lo || hi != tc.hi || n != tc.next) {
				t.Fatalf("%d %d %d", lo, hi, n)
			}
		})
	}
	// Exhaustively compare the linear formula with all possible affix pairs on a
	// small alphabet. This tests interval identity, not implementation's report.
	words := []string{""}
	for n := 1; n <= 4; n++ {
		for mask := 0; mask < 1<<n; mask++ {
			s := make([]byte, n)
			for i := range s {
				s[i] = 'a' + byte(mask>>i&1)
			}
			words = append(words, string(s))
		}
	}
	for _, a := range words {
		for _, b := range words {
			if a == b {
				continue
			}
			best := -1
			pairs := [][2]int{}
			for p := 0; p <= min(len(a), len(b)); p++ {
				if a[:p] != b[:p] {
					continue
				}
				for s := 0; s <= min(len(a)-p, len(b)-p); s++ {
					if a[len(a)-s:] != b[len(b)-s:] {
						continue
					}
					if p+s > best {
						best = p + s
						pairs = nil
					}
					if p+s == best {
						pairs = append(pairs, [2]int{p, s})
					}
				}
			}
			lo, hi, next, err := changedInterval([]rune(a), []rune(b))
			if (len(pairs) != 1) != (err != nil) {
				t.Fatalf("%q->%q pairs=%v err=%v", a, b, pairs, err)
			}
			if err == nil && (lo != pairs[0][0] || hi != len(a)-pairs[0][1] || next != len(b)-pairs[0][1]) {
				t.Fatalf("wrong localization %q->%q", a, b)
			}
		}
	}
}
