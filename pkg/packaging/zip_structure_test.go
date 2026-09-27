package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"testing"
)

func TestZIPStructureContracts(t *testing.T) {
	p := New()
	_, _ = p.AddPart("a.xml", ContentTypeXML, []byte(`<data/>`))
	var b bytes.Buffer
	if err := p.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		name   string
		mutate func([]byte)
	}{
		{"name", func(d []byte) { d[30] ^= 1 }},
		{"method", func(d []byte) { binary.LittleEndian.PutUint16(d[8:10], zip.Store) }},
		{"flags", func(d []byte) { d[6] ^= 1 }},
		{"central offset points at another header", func(d []byte) {
			first := bytes.Index(d, []byte{'P', 'K', 1, 2})
			second := bytes.Index(d[first+4:], []byte{'P', 'K', 1, 2}) + first + 4
			binary.LittleEndian.PutUint32(d[second+42:second+46], 0)
		}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			d := bytes.Clone(b.Bytes())
			tc.mutate(d)
			q, err := OpenBytes(d)
			if q != nil {
				_ = q.Close()
			}
			if err == nil {
				t.Fatal("inconsistent archive accepted")
			}
		})
	}
}
