package losslessxml

import "testing"

func TestQNameAndBoundaryBatch(t *testing.T) {
	for _, bad := range []string{`<p:1 xmlns:p="u"/>`, `<r xmlns:p="u" p:1="x"/>`, `<r xmlns:1="u"/>`, `<r xmlns:p="u"><p:́/></r>`, "\u00a0<r/>", "<r/>\u2003", "\ufeff\ufeff<r/>"} {
		t.Run(bad, func(t *testing.T) {
			if _, err := Parse([]byte(bad)); err == nil {
				t.Fatal("invalid QName/boundary accepted")
			}
		})
	}
	for _, good := range []string{`<r xmlns:p="u"><p:_name/></r>`, `<r xmlns:p="u"><p:名/></r>`, `<r a·b="1"/>`, `<r xmlns:p="u"><p:á/></r>`, " \t\r\n<r/>\n"} {
		t.Run(good, func(t *testing.T) {
			if _, err := Parse([]byte(good)); err != nil {
				t.Fatal(err)
			}
		})
	}
}
