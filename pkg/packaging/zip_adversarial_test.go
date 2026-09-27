package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"errors"
	"strings"
	"testing"
)

func TestZIPDescriptorAndPrefixBatch(t *testing.T) {
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	w, err := z.Create("data.bin")
	if err != nil {
		t.Fatal(err)
	}
	_, _ = w.Write([]byte("payload"))
	if err = z.Close(); err != nil {
		t.Fatal(err)
	}
	source := b.Bytes()
	t.Run("absolute adjusted prefixed archive refuses", func(t *testing.T) {
		prefix := []byte("PREPENDED")
		d := append(bytes.Clone(prefix), source...)
		central := bytes.Index(d, []byte{'P', 'K', 1, 2})
		end := bytes.LastIndex(d, []byte{'P', 'K', 5, 6})
		old := binary.LittleEndian.Uint32(d[central+42 : central+46])
		binary.LittleEndian.PutUint32(d[central+42:central+46], old+uint32(len(prefix)))
		binary.LittleEndian.PutUint32(d[end+16:end+20], uint32(central))
		r, err := zip.NewReader(bytes.NewReader(d), int64(len(d)))
		if err != nil {
			t.Fatal(err)
		}
		if err = validateZIPStructure(bytes.NewReader(d), int64(len(d)), r.File); err == nil {
			t.Fatal("adjusted prefix accepted")
		}
	})
	t.Run("unsigned descriptor whose CRC equals signature", func(t *testing.T) {
		r, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
		if err != nil {
			t.Fatal(err)
		}
		start, err := r.File[0].DataOffset()
		if err != nil {
			t.Fatal(err)
		}
		descriptor := int(start + int64(r.File[0].CompressedSize64))
		if little.Uint32(source[descriptor:descriptor+4]) != 0x08074b50 {
			t.Fatal("expected signed fixture")
		}
		d := append(bytes.Clone(source[:descriptor]), source[descriptor+4:]...)
		central := bytes.Index(d, []byte{'P', 'K', 1, 2})
		end := bytes.LastIndex(d, []byte{'P', 'K', 5, 6})
		little.PutUint32(d[descriptor:descriptor+4], 0x08074b50)
		little.PutUint32(d[central+16:central+20], 0x08074b50)
		little.PutUint32(d[end+16:end+20], uint32(central))
		// This native construction is independent of the canonical generator.
		// It checks the same geometry and payload refusal without treating a
		// generated ZIP's entire byte sequence as an acceptance oracle.
		original := bytes.Clone(d)
		r, err = zip.NewReader(bytes.NewReader(d), int64(len(d)))
		if err != nil || len(r.File) != 1 {
			t.Fatalf("one-member archive: %v", err)
		}
		if r.File[0].Flags&8 == 0 || r.File[0].Method != zip.Deflate || r.File[0].UncompressedSize64 != 7 || r.File[0].CRC32 != 0x08074b50 || descriptor+12 != central || little.Uint32(d[central:]) != 0x02014b50 {
			t.Fatal("unsigned 12-byte descriptor geometry drift")
		}
		if err = validateZIPStructure(bytes.NewReader(d), int64(len(d)), r.File); err != nil {
			t.Fatalf("legal unsigned descriptor geometry rejected: %v", err)
		}
		// Payload CRC is deliberately not recalculated: structural ambiguity is
		// resolved before full intake independently refuses the bad payload.
		q, err := OpenReaderWithLimits(bytes.NewReader(d), int64(len(d)), Limits{})
		var refusal *Refusal
		if q != nil || !errors.As(err, &refusal) || refusal.Kind != "invalid_package" || refusal.Operation != "open" || !strings.Contains(strings.ToLower(refusal.Detail), "checksum") || strings.Contains(strings.ToLower(refusal.Detail), "descriptor") {
			t.Fatalf("wrong payload CRC refusal: package=%v err=%v", q, err)
		}
		if !bytes.Equal(d, original) {
			t.Fatal("caller bytes altered")
		}
	})
}
