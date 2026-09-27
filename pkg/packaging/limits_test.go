package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"strings"
	"testing"
)

func TestIntakeBudgetAndIntegrityContracts(t *testing.T) {
	p := New()
	_, _ = p.AddPart("data.xml", ContentTypeXML, []byte(`<data/>`))
	var b bytes.Buffer
	if err := p.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	data := b.Bytes()
	for _, tc := range []struct {
		name  string
		limit Limits
	}{
		{"source", Limits{MaxSourceBytes: 1}}, {"entries", Limits{MaxEntries: 1}}, {"part", Limits{MaxPartBytes: 1}}, {"total", Limits{MaxTotalBytes: 1}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			q, err := OpenReaderWithLimits(bytes.NewReader(data), int64(len(data)), tc.limit)
			if q != nil {
				_ = q.Close()
			}
			var refusal *Refusal
			if !errors.As(err, &refusal) || refusal.Kind != "resource_limit" {
				t.Fatalf("got %v", err)
			}
		})
	}
	t.Run("invalid arguments", func(t *testing.T) {
		for _, limit := range []Limits{{MaxSourceBytes: -1}, {MaxEntries: -1}} {
			if _, err := OpenReaderWithLimits(bytes.NewReader(data), int64(len(data)), limit); err == nil {
				t.Fatal("negative budget accepted")
			}
		}
	})
	t.Run("CRC typed refusal", func(t *testing.T) {
		var b bytes.Buffer
		z := zip.NewWriter(&b)
		w, err := z.CreateHeader(&zip.FileHeader{Name: "payload.bin", Method: zip.Store})
		if err != nil {
			t.Fatal(err)
		}
		_, _ = w.Write([]byte("UNIQUE_PAYLOAD"))
		if err = z.Close(); err != nil {
			t.Fatal(err)
		}
		corrupt := bytes.Clone(b.Bytes())
		at := bytes.Index(corrupt, []byte("UNIQUE_PAYLOAD"))
		if at < 0 {
			t.Fatal("payload not found")
		}
		corrupt[at] ^= 1
		_, err = OpenBytes(corrupt)
		var refusal *Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "invalid_package" {
			t.Fatalf("got %v", err)
		}
	})
	t.Run("directory and empty binary", func(t *testing.T) {
		var b bytes.Buffer
		z := zip.NewWriter(&b)
		entries := []struct{ name, text string }{{"assets/", ""}, {"[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/></Types>`}, {"assets/empty.bin", ""}}
		for _, e := range entries {
			w, err := z.Create(e.name)
			if err != nil {
				t.Fatal(err)
			}
			_, _ = w.Write([]byte(e.text))
		}
		if err := z.Close(); err != nil {
			t.Fatal(err)
		}
		q, err := OpenReaderWithLimits(bytes.NewReader(b.Bytes()), int64(b.Len()), Limits{MaxEntries: 3, MaxPartBytes: 1024, MaxTotalBytes: 1024})
		if err != nil {
			t.Fatal(err)
		}
		defer q.Close()
		for _, p := range q.Parts() {
			if strings.HasSuffix(p.URI(), "/") {
				t.Fatal("directory treated as part")
			}
		}
	})
}
