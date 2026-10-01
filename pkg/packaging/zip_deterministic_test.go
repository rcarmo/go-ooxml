package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"strings"
	"testing"
)

func TestWriteZIP32DeterministicFamily(t *testing.T) {
	binary := make([]byte, 256)
	for i := range binary {
		binary[i] = byte(i)
	}
	entries := []ZIP32Entry{{Name: "[Content_Types].xml", Data: []byte(`<Types/>`)}, {Name: "custom/data.bin", Data: binary}, {Name: "word/document.xml", Data: []byte(`<w:document>` + strings.Repeat("A", 2048) + `</w:document>`)}}
	originals := make([][]byte, len(entries))
	for i, e := range entries {
		originals[i] = bytes.Clone(e.Data)
	}
	first, err := WriteZIP32Deterministic(entries)
	if err != nil {
		t.Fatal(err)
	}
	second, err := WriteZIP32Deterministic(entries)
	if err != nil || !bytes.Equal(first, second) {
		t.Fatalf("writer not deterministic: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(first), int64(len(first)))
	if err != nil {
		t.Fatal(err)
	}
	methods := map[uint16]bool{}
	if len(zr.File) != len(entries) {
		t.Fatalf("entries %d", len(zr.File))
	}
	for i, f := range zr.File {
		if f.Name != entries[i].Name || f.Flags&0x800 == 0 {
			t.Fatalf("order/UTF8 flag %d: %q %x", i, f.Name, f.Flags)
		}
		methods[f.Method] = true
		in, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		var b bytes.Buffer
		_, err = b.ReadFrom(in)
		closeErr := in.Close()
		if err != nil || closeErr != nil || !bytes.Equal(b.Bytes(), originals[i]) {
			t.Fatalf("reopened payload %d: %v %v", i, err, closeErr)
		}
		if !bytes.Equal(entries[i].Data, originals[i]) {
			t.Fatal("caller entry changed")
		}
	}
	if !methods[zip.Store] || !methods[zip.Deflate] {
		t.Fatalf("methods: %v", methods)
	}
	duplicate := []ZIP32Entry{{Name: "A.bin", Data: []byte("a")}, {Name: "a.bin", Data: []byte("b")}}
	result, err := WriteZIP32Deterministic(duplicate)
	var refusal *ProfileError
	if result != nil || !errors.As(err, &refusal) || refusal.Reason != "zip-case-collision" {
		t.Fatalf("collision: %v", err)
	}
	if string(duplicate[0].Data) != "a" || string(duplicate[1].Data) != "b" {
		t.Fatal("refusal changed caller entries")
	}
	// Count sentinel would cause archive/zip to emit a ZIP64 end record.
	tooMany := make([]ZIP32Entry, 65535)
	for i := range tooMany {
		tooMany[i].Name = fmt.Sprintf("p/%05d", i)
	}
	if result, err := WriteZIP32Deterministic(tooMany); result != nil || !errors.As(err, &refusal) || refusal.Reason != "zip-structure-invalid" {
		t.Fatalf("ZIP64 entry sentinel escaped ZIP32 profile: %v", err)
	}
	if tooMany[0].Name != "p/00000" || tooMany[len(tooMany)-1].Name != "p/65534" {
		t.Fatal("refusal changed caller entries")
	}
}
