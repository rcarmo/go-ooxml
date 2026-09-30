package packaging_test

import (
	"bytes"
	"compress/flate"
	"encoding/binary"
	"errors"
	"hash/crc32"
	"math"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type zipProfileSpec struct {
	name, localName  string
	payload          []byte
	method, flags    uint16
	crc              uint32
	size, compressed int
}

// A local generator deliberately does not use the production writer. All
// headers and end-record offsets are explicit and can be corrupted separately.
func zipProfileSample(specs []zipProfileSpec) []byte {
	var body, directory bytes.Buffer
	put16 := func(b *bytes.Buffer, n uint16) { _ = binary.Write(b, binary.LittleEndian, n) }
	put32 := func(b *bytes.Buffer, n uint32) { _ = binary.Write(b, binary.LittleEndian, n) }
	for _, s := range specs {
		localName := s.localName
		if localName == "" {
			localName = s.name
		}
		method := s.method
		if method == 0 {
			method = 8
		}
		var packed bytes.Buffer
		if method == 8 {
			z, _ := flate.NewWriter(&packed, flate.DefaultCompression)
			_, _ = z.Write(s.payload)
			_ = z.Close()
		} else {
			packed.Write(s.payload)
		}
		compressed := packed.Len()
		if s.compressed > 0 {
			compressed = s.compressed
		}
		size := len(s.payload)
		if s.size > 0 {
			size = s.size
		}
		crc := crc32.ChecksumIEEE(s.payload)
		if s.crc != 0 {
			crc = s.crc
		}
		offset := body.Len()
		put32(&body, 0x04034b50)
		put16(&body, 20)
		put16(&body, s.flags|0x800)
		put16(&body, method)
		put16(&body, 0)
		put16(&body, 0)
		put32(&body, crc)
		put32(&body, uint32(compressed))
		put32(&body, uint32(size))
		put16(&body, uint16(len(localName)))
		put16(&body, 0)
		body.WriteString(localName)
		body.Write(packed.Bytes())
		put32(&directory, 0x02014b50)
		put16(&directory, 20)
		put16(&directory, 20)
		put16(&directory, s.flags|0x800)
		put16(&directory, method)
		put16(&directory, 0)
		put16(&directory, 0)
		put32(&directory, crc)
		put32(&directory, uint32(compressed))
		put32(&directory, uint32(size))
		put16(&directory, uint16(len(s.name)))
		put16(&directory, 0)
		put16(&directory, 0)
		put16(&directory, 0)
		put16(&directory, 0)
		put32(&directory, 0)
		put32(&directory, uint32(offset))
		directory.WriteString(s.name)
	}
	result := append([]byte{}, body.Bytes()...)
	centralAt := len(result)
	result = append(result, directory.Bytes()...)
	var end bytes.Buffer
	put32(&end, 0x06054b50)
	put16(&end, 0)
	put16(&end, 0)
	put16(&end, uint16(len(specs)))
	put16(&end, uint16(len(specs)))
	put32(&end, uint32(directory.Len()))
	put32(&end, uint32(centralAt))
	put16(&end, 0)
	return append(result, end.Bytes()...)
}

func zipProfileReason(t *testing.T, got []packaging.ZIP32Entry, err error, reason string) {
	t.Helper()
	var refusal *packaging.ProfileError
	if got != nil || !errors.As(err, &refusal) || refusal.Reason != reason {
		t.Fatalf("output=%v refusal=%v, want %s", got, err, reason)
	}
}

func TestRawZIP32ProfileRefusals(t *testing.T) {
	one := []zipProfileSpec{{name: "word/document.xml", payload: []byte("payload")}}
	for _, tc := range []struct {
		name   string
		specs  []zipProfileSpec
		mutate func([]byte)
		reason string
	}{
		{"duplicate", []zipProfileSpec{{name: "word/document.xml", payload: []byte("one")}, {name: "word/document.xml", payload: []byte("two")}}, nil, "zip-duplicate-entry"},
		{"ASCII case", []zipProfileSpec{{name: "word/document.xml", payload: []byte("one")}, {name: "WORD/document.xml", payload: []byte("two")}}, nil, "zip-case-collision"},
		{"parent traversal", []zipProfileSpec{{name: "../word/document.xml", payload: []byte("bad")}}, nil, "zip-name-invalid"},
		{"encryption", []zipProfileSpec{{name: "word/document.xml", payload: []byte("secret"), flags: 1}}, nil, "zip-encryption-unsupported"},
		{"unsupported method", one, func(d []byte) {
			binary.LittleEndian.PutUint16(d[8:10], 12)
			at := bytes.Index(d, []byte{'P', 'K', 1, 2})
			binary.LittleEndian.PutUint16(d[at+10:at+12], 12)
		}, "zip-method-unsupported"},
		{"disk", one, func(d []byte) { binary.LittleEndian.PutUint16(d[len(d)-18:], 1) }, "zip-multi-disk-unsupported"},
		{"missing ZIP64", one, func(d []byte) {
			binary.LittleEndian.PutUint16(d[len(d)-14:], 65535)
			binary.LittleEndian.PutUint16(d[len(d)-12:], 65535)
		}, "zip-structure-invalid"},
		{"local name mismatch", []zipProfileSpec{{name: "word/document.xml", localName: "word/other.xml", payload: []byte("x")}}, nil, "zip-local-metadata-mismatch"},
		{"CRC mismatch", []zipProfileSpec{{name: "word/document.xml", payload: []byte("payload"), crc: 0xdeadbeef}}, nil, "zip-crc-mismatch"},
		{"stored size mismatch", one, func(d []byte) {
			binary.LittleEndian.PutUint16(d[8:10], 0)
			at := bytes.Index(d, []byte{'P', 'K', 1, 2})
			binary.LittleEndian.PutUint16(d[at+10:at+12], 0)
			for _, off := range []int{18, 22} {
				binary.LittleEndian.PutUint32(d[off:off+4], 99)
			}
			for _, off := range []int{at + 20, at + 24} {
				binary.LittleEndian.PutUint32(d[off:off+4], 99)
			}
		}, "zip-size-mismatch"},
		{"deflate overrun", []zipProfileSpec{{name: "word/document.xml", payload: []byte(strings.Repeat("A", 4096)), size: 32}}, nil, "zip-size-mismatch"},
		{"trailing byte", one, func(d []byte) {}, "zip-end-record-missing"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data := zipProfileSample(tc.specs)
			if tc.name == "trailing byte" {
				data = append(data, '\n')
			} else if tc.mutate != nil {
				tc.mutate(data)
			}
			original := bytes.Clone(data)
			got, err := packaging.ReadZIP32(data, packaging.ZIP32Limits{})
			zipProfileReason(t, got, err, tc.reason)
			if !bytes.Equal(data, original) {
				t.Fatal("caller ZIP mutated")
			}
		})
	}
}

func TestRawZIP32FiniteDefaultsAndASCIICollisionPolicy(t *testing.T) {
	// The empty limits struct must not silently disable every budget.
	oversized := make([]byte, (64<<20)+1)
	got, err := packaging.ReadZIP32(oversized, packaging.ZIP32Limits{})
	zipProfileReason(t, got, err, "zip-archive-too-large")
	for _, ratio := range []float64{math.NaN(), math.Inf(1), math.Inf(-1), -1} {
		got, err := packaging.ReadZIP32([]byte("valid or invalid"), packaging.ZIP32Limits{MaxCompressionRatio: ratio})
		if got != nil || err == nil {
			t.Fatalf("nonfinite/negative ratio accepted: %v", ratio)
		}
		var refusal *packaging.ProfileError
		if errors.As(err, &refusal) {
			t.Fatalf("invalid programmer budget classified as ZIP sample failure: %v", err)
		}
	}
	for _, tc := range []struct {
		limit  packaging.ZIP32Limits
		reason string
	}{
		{packaging.ZIP32Limits{MaxEntries: 1}, "zip-too-many-entries"},
		{packaging.ZIP32Limits{MaxEntryBytes: 8}, "zip-entry-too-large"},
		{packaging.ZIP32Limits{MaxTotalBytes: 8}, "zip-total-too-large"},
		{packaging.ZIP32Limits{MaxCompressionRatio: 2}, "zip-compression-ratio-exceeded"},
	} {
		data := zipProfileSample([]zipProfileSpec{{name: "a.bin", payload: []byte(strings.Repeat("A", 4096))}, {name: "b.bin", payload: []byte("bb")}})
		got, err := packaging.ReadZIP32(data, tc.limit)
		zipProfileReason(t, got, err, tc.reason)
	}
	// Unicode case pairs differ in their raw UTF-8 bytes. Only ASCII letters
	// fold; both entries must survive writer and reader with distinct payloads.
	inputs := []packaging.ZIP32Entry{{Name: "word/Å.xml", Data: []byte("upper")}, {Name: "word/å.xml", Data: []byte("lower")}}
	data, err := packaging.WriteZIP32(inputs)
	if err != nil {
		t.Fatal(err)
	}
	out, err := packaging.ReadZIP32(data, packaging.ZIP32Limits{})
	if err != nil || !reflect.DeepEqual(out, inputs) {
		t.Fatalf("Unicode aliases were folded: %v %v", out, err)
	}
	control := zipProfileSample([]zipProfileSpec{{name: "word/Å.xml", payload: []byte("upper")}, {name: "word/å.xml", payload: []byte("lower")}})
	out, err = packaging.ReadZIP32(control, packaging.ZIP32Limits{})
	if err != nil || !reflect.DeepEqual(out, inputs) {
		t.Fatalf("independent Unicode pair: %v %v", out, err)
	}
}

func TestRawZIP32ProfileBoundsAndWriter(t *testing.T) {
	a := []byte(strings.Repeat("A", 4096))
	data := zipProfileSample([]zipProfileSpec{{name: "a.bin", payload: a}, {name: "b.bin", payload: []byte("bb")}})
	control, err := packaging.ReadZIP32(data, packaging.ZIP32Limits{})
	if err != nil || len(control) != 2 || control[0].Name != "a.bin" || !bytes.Equal(control[0].Data, a) || control[1].Name != "b.bin" || string(control[1].Data) != "bb" {
		t.Fatalf("control=%v: %v", control, err)
	}
	for _, tc := range []struct {
		name   string
		limit  packaging.ZIP32Limits
		reason string
	}{
		{"archive", packaging.ZIP32Limits{MaxArchiveBytes: int64(len(data) - 1)}, "zip-archive-too-large"},
		{"entries", packaging.ZIP32Limits{MaxEntries: 1}, "zip-too-many-entries"},
		{"entry", packaging.ZIP32Limits{MaxEntryBytes: 8}, "zip-entry-too-large"},
		{"total", packaging.ZIP32Limits{MaxTotalBytes: 8}, "zip-total-too-large"},
		{"ratio", packaging.ZIP32Limits{MaxCompressionRatio: 2}, "zip-compression-ratio-exceeded"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			before := bytes.Clone(data)
			got, err := packaging.ReadZIP32(data, tc.limit)
			zipProfileReason(t, got, err, tc.reason)
			if !bytes.Equal(data, before) {
				t.Fatal("caller ZIP mutated")
			}
		})
	}
	for _, tc := range []struct {
		name    string
		entries []packaging.ZIP32Entry
		reason  string
	}{
		{"case collision", []packaging.ZIP32Entry{{Name: "word/document.xml", Data: []byte("one")}, {Name: "WORD/document.xml", Data: []byte("two")}}, "zip-case-collision"},
		{"directory payload", []packaging.ZIP32Entry{{Name: "word/", Data: []byte("not empty")}}, "zip-directory-entry-invalid"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			original := make([]packaging.ZIP32Entry, len(tc.entries))
			for i, e := range tc.entries {
				original[i] = packaging.ZIP32Entry{Name: e.Name, Data: bytes.Clone(e.Data)}
			}
			got, err := packaging.WriteZIP32(tc.entries)
			if got != nil {
				t.Fatalf("writer returned partial archive: %x", got)
			}
			var refusal *packaging.ProfileError
			if !errors.As(err, &refusal) || refusal.Reason != tc.reason {
				t.Fatalf("writer: %v", err)
			}
			if !reflect.DeepEqual(tc.entries, original) {
				t.Fatal("writer modified input")
			}
		})
	}
	written, err := packaging.WriteZIP32([]packaging.ZIP32Entry{{Name: "a.bin", Data: []byte("one")}, {Name: "b.bin", Data: []byte("two")}})
	if err != nil {
		t.Fatal(err)
	}
	reopened, err := packaging.ReadZIP32(written, packaging.ZIP32Limits{})
	if err != nil || len(reopened) != 2 || string(reopened[0].Data) != "one" || string(reopened[1].Data) != "two" {
		t.Fatalf("writer reopen: %v: %v", reopened, err)
	}
}
