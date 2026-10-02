package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"strings"
	"testing"
)

type budgetReadProbe struct {
	data  []byte
	reads int
}

func (p *budgetReadProbe) ReadAt(dst []byte, offset int64) (int, error) {
	p.reads++
	return bytes.NewReader(p.data).ReadAt(dst, offset)
}

func TestNegativeBudgetReadBoundary(t *testing.T) {
	pkg := New()
	if _, err := pkg.AddPart("data.xml", ContentTypeXML, []byte(`<data/>`)); err != nil {
		t.Fatal(err)
	}
	var buf bytes.Buffer
	if err := pkg.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	source := bytes.Clone(buf.Bytes())
	for _, tc := range []struct {
		name   string
		size   int64
		limits Limits
	}{
		{"source bytes", int64(len(source)), Limits{MaxSourceBytes: -1}},
		{"entry count", int64(len(source)), Limits{MaxEntries: -1}},
		{"archive size", -1, Limits{}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			probe := &budgetReadProbe{data: source}
			opened, err := OpenReaderWithLimits(probe, tc.size, tc.limits)
			var refused *Refusal
			if opened != nil || err == nil || err.Error() != "negative archive size or resource budget" || errors.As(err, &refused) || probe.reads != 0 || !bytes.Equal(source, probe.data) {
				t.Fatalf("invalid input: package=%v error=%v attempted reads=%d", opened, err, probe.reads)
			}
		})
	}
	for _, tc := range []struct {
		name   string
		limits Limits
	}{
		{"default budgets", Limits{}},
		{"explicit zero budgets", Limits{MaxSourceBytes: 0, MaxEntries: 0}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			probe := &budgetReadProbe{data: source}
			opened, err := OpenReaderWithLimits(probe, int64(len(source)), tc.limits)
			if err != nil || opened == nil {
				t.Fatalf("valid input refused: %v", err)
			}
			defer opened.Close()
			if probe.reads == 0 || len(opened.Parts()) == 0 || !bytes.Equal(source, probe.data) {
				t.Fatalf("valid input not read or changed: reads=%d", probe.reads)
			}
		})
	}
}

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
