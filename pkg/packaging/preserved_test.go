package packaging

import (
	"bytes"
	"errors"
	"io"
	"testing"
)

func TestRetainedSourceContracts(t *testing.T) {
	makeSession := func() (*Preserved, []byte) {
		t.Helper()
		q := New()
		_, _ = q.AddPart("a.xml", ContentTypeXML, []byte(`<a/>`))
		_, _ = q.AddPart("b.xml", ContentTypeXML, []byte(`<b/>`))
		var b bytes.Buffer
		if err := q.WriteTo(&b); err != nil {
			t.Fatal(err)
		}
		source := bytes.Clone(b.Bytes())
		p, err := OpenPreserved(source, Limits{})
		if err != nil {
			t.Fatal(err)
		}
		return p, source
	}
	t.Run("source and part isolation", func(t *testing.T) {
		p, source := makeSession()
		original := bytes.Clone(source)
		source[0] ^= 1
		part, hash, err := p.Part("a.xml")
		if err != nil {
			t.Fatal(err)
		}
		if string(part) != `<a/>` || hash != fingerprint(part) {
			t.Fatal("missing payload or fingerprint")
		}
		part[0] = '!'
		var b bytes.Buffer
		if err = p.WriteTo(&b); err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(original, b.Bytes()) {
			t.Fatal("caller data mutated retained package")
		}
	})
	t.Run("atomic mixed batch", func(t *testing.T) {
		p, source := makeSession()
		_, h, _ := p.Part("a.xml")
		err := p.Replace([]Replacement{{"a.xml", h, []byte(`<new/>`)}, {"b.xml", "stale", []byte(`<new/>`)}})
		var r *Refusal
		if !errors.As(err, &r) {
			t.Fatalf("got %v", err)
		}
		var b bytes.Buffer
		_ = p.WriteTo(&b)
		if !bytes.Equal(source, b.Bytes()) {
			t.Fatal("partial mutation")
		}
	})
	t.Run("malformed XML refuses", func(t *testing.T) {
		p, _ := makeSession()
		_, h, _ := p.Part("a.xml")
		if err := p.Replace([]Replacement{{"a.xml", h, []byte(`<a>`)}}); err == nil {
			t.Fatal("invalid XML accepted")
		}
	})
	t.Run("registry changes refuse", func(t *testing.T) {
		p, _ := makeSession()
		b, h, _ := p.Part(ContentTypesPath)
		if err := p.Replace([]Replacement{{ContentTypesPath, h, b}}); err == nil {
			t.Fatal("registry replacement accepted")
		}
	})
	t.Run("duplicate targets refuse", func(t *testing.T) {
		p, _ := makeSession()
		_, h, _ := p.Part("a.xml")
		if err := p.Replace([]Replacement{{"a.xml", h, []byte(`<n/>`)}, {"a.xml", h, []byte(`<m/>`)}}); err == nil {
			t.Fatal("duplicate replacements accepted")
		}
	})
	t.Run("short no-op writer", func(t *testing.T) {
		p, _ := makeSession()
		if err := p.WriteTo(shortWriter{}); !errors.Is(err, io.ErrShortWrite) {
			t.Fatalf("got %v", err)
		}
	})
}

type shortWriter struct{}

func (shortWriter) Write(b []byte) (int, error) { return len(b) - 1, nil }
