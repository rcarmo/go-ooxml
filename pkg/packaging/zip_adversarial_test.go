package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
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
		r, err = zip.NewReader(bytes.NewReader(d), int64(len(d)))
		if err != nil {
			t.Fatal(err)
		}
		if err = validateZIPStructure(bytes.NewReader(d), int64(len(d)), r.File); err != nil {
			t.Fatalf("legal unsigned descriptor geometry rejected: %v", err)
		}
		// Payload CRC is deliberately not recalculated: this test isolates structural
		// ambiguity. Full OpenBytes must still reject the mismatched payload CRC.
		if _, err = OpenBytes(d); err == nil {
			t.Fatal("bad payload CRC accepted")
		}
	})
}
