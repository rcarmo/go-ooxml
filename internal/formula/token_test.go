package formula

import "testing"

func TestLiteralPunctuationBatch(t *testing.T) {
	for _, source := range []string{`A1 "+" B1`, `A1 ":" B2`, `A1 "!" B1`, `A1 "%"`, `"=" A1`} {
		t.Run(source, func(t *testing.T) {
			if refs, err := Analyze(source); err == nil || len(refs) != 0 {
				t.Fatalf("invalid accepted %v %v", refs, err)
			}
		})
	}
	for _, source := range []string{`SUM(")",A1)`, `IF(A1=1,")",B1)`, `"("&A1`, `SUM("%",A1)`, `"="`} {
		t.Run(source, func(t *testing.T) {
			if _, err := Analyze(source); err != nil {
				t.Fatal(err)
			}
		})
	}
}
