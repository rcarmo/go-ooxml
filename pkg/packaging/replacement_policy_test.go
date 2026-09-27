package packaging

import (
	"bytes"
	"errors"
	"testing"
)

func TestReplacementXMLNamespacePolicy(t *testing.T) {
	for _, text := range []string{`<x:r/>`, `<r x:a="1"/>`, `<r xmlns:a="u" xmlns:b="u" a:n="1" b:n="2"/>`} {
		t.Run(text, func(t *testing.T) {
			q := New()
			_, _ = q.AddPart("a.xml", ContentTypeXML, []byte(`<a/>`))
			var b bytes.Buffer
			if err := q.WriteTo(&b); err != nil {
				t.Fatal(err)
			}
			p, err := OpenPreserved(b.Bytes(), Limits{})
			if err != nil {
				t.Fatal(err)
			}
			_, hash, _ := p.Part("a.xml")
			err = p.Replace([]Replacement{{Part: "a.xml", ExpectedSHA256: hash, Data: []byte(text)}})
			var refusal *Refusal
			if !errors.As(err, &refusal) {
				t.Fatalf("expected namespace refusal, got %v", err)
			}
		})
	}
}
