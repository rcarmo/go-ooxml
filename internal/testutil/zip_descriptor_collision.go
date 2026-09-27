package testutil

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"fmt"
)

// UnsignedDescriptorCollisionZIP has a seven-byte DEFLATED data.bin payload.
// Its unsigned descriptor CRC is 0x08074b50, equal to the optional signature,
// while the actual payload CRC is different. The archive's geometry is valid.
func UnsignedDescriptorCollisionZIP() ([]byte, error) {
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	w, err := zw.Create("data.bin")
	if err != nil {
		return nil, err
	}
	if _, err = w.Write([]byte("payload")); err != nil {
		return nil, err
	}
	if err = zw.Close(); err != nil {
		return nil, err
	}
	source := out.Bytes()
	zr, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil || len(zr.File) != 1 {
		return nil, fmt.Errorf("unexpected ZIP writer output: %v", err)
	}
	offset, err := zr.File[0].DataOffset()
	if err != nil {
		return nil, err
	}
	at := int(offset + int64(zr.File[0].CompressedSize64))
	le := binary.LittleEndian
	if at < 0 || at+16 > len(source) || le.Uint32(source[at:]) != 0x08074b50 {
		return nil, fmt.Errorf("expected signed writer descriptor")
	}
	archive := append(bytes.Clone(source[:at]), source[at+4:]...)
	central := bytes.Index(archive, []byte{'P', 'K', 1, 2})
	end := bytes.LastIndex(archive, []byte{'P', 'K', 5, 6})
	if central < at+12 || end < central+46 || end+22 != len(archive) {
		return nil, fmt.Errorf("unexpected ZIP directory bounds")
	}
	le.PutUint32(archive[at:], 0x08074b50)
	le.PutUint32(archive[central+16:], 0x08074b50)
	le.PutUint32(archive[end+16:], uint32(central))
	return archive, nil
}
