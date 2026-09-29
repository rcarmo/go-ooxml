package losslessxml

import (
	"bytes"
	"strings"
	"testing"
)

func TestEquivalentControls(t *testing.T) {
	for _, tc := range []struct {
		name, left, right string
		want              bool
	}{
		{"safe identical", `<a><b>x</b></a>`, `<a><b>x</b></a>`, true},
		{"prefix alias", `<p:a xmlns:p="urn:one" p:x="1"/>`, `<q:a xmlns:q="urn:one" q:x="1"/>`, true},
		{"attribute order", `<a x="1" y="2"/>`, `<a y="2" x="1"/>`, true},
		{"different child order", `<a><b/><c/></a>`, `<a><c/><b/></a>`, false},
		{"different whitespace", `<a> x </a>`, `<a>x</a>`, false},
		{"different attr value", `<a x="1"/>`, `<a x="2"/>`, false},
		{"prefix attribute binding", `<a xmlns:p="urn:one" value="p:x"/>`, `<a xmlns:p="urn:two" value="p:x"/>`, false},
		{"ordinary colon text spelling", `<a xmlns:p="urn:one" xmlns:q="urn:one">p:x</a>`, `<a xmlns:p="urn:one" xmlns:q="urn:one">q:x</a>`, false},
		{"ordinary colon attribute spelling", `<a xmlns:p="urn:one" xmlns:q="urn:one" value="p:x"/>`, `<a xmlns:p="urn:one" xmlns:q="urn:one" value="q:x"/>`, false},
		{"PI separator", `<?p  x?><a/>`, `<?p x?><a/>`, false},
		{"PI data", `<?p old?><a/>`, `<?p new?><a/>`, false},
		{"prefix alias with PI", `<?p x?><p:a xmlns:p="urn:one"/>`, `<?p x?><q:a xmlns:q="urn:one"/>`, true},
		{"ordered PI", `<a><?p x?></a>`, `<a><?q x?></a>`, false},
		{"prolog comment", `<!--old--><a/>`, `<!--new--><a/>`, false},
		{"DTD identical", `<!DOCTYPE a><a/>`, `<!DOCTYPE a><a/>`, false},
		{"malformed identical", `broken`, `broken`, false},
		{"unbound prefix", `<a value="p:x"/>`, `<a value="p:x"/>`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			left, right := []byte(tc.left), []byte(tc.right)
			l0, r0 := bytes.Clone(left), bytes.Clone(right)
			got := Equivalent(left, right)
			if got != tc.want || !bytes.Equal(left, l0) || !bytes.Equal(right, r0) {
				t.Fatalf("result=%t want=%t or operands changed", got, tc.want)
			}
		})
	}
	for _, input := range [][]byte{bytes.Repeat([]byte{'a'}, (64<<10)+1), []byte{0xff}, []byte(strings.Repeat("<a>", 513) + strings.Repeat("</a>", 513))} {
		left, right := bytes.Clone(input), bytes.Clone(input)
		if Equivalent(left, right) || !bytes.Equal(left, input) || !bytes.Equal(right, input) {
			t.Fatal("unsafe or oversized identical input accepted, or bytes changed")
		}
	}
}
