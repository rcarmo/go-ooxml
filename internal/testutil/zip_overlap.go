package testutil

import (
	"bytes"
	"encoding/binary"
	"hash/crc32"
)

// OverlappingZIP contains distinct stored entries with coherent local/central
// names, sizes and CRCs. The outer payload contains the inner local header and
// payload. Both entries read successfully in archive/zip but physically overlap.
func OverlappingZIP() []byte {
	le := binary.LittleEndian
	local := func(name string, payload []byte) []byte {
		h := make([]byte, 30)
		le.PutUint32(h, 0x04034b50)
		le.PutUint16(h[4:], 20)
		le.PutUint32(h[14:], crc32.ChecksumIEEE(payload))
		le.PutUint32(h[18:], uint32(len(payload)))
		le.PutUint32(h[22:], uint32(len(payload)))
		le.PutUint16(h[26:], uint16(len(name)))
		h = append(h, []byte(name)...)
		return append(h, payload...)
	}
	ct := []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/></Types>`)
	inner := []byte("inner data")
	outer := local("inner.bin", inner)
	ctRecord := local("[Content_Types].xml", ct)
	outerRecord := local("outer.bin", outer)
	var out bytes.Buffer
	out.Write(ctRecord)
	out.Write(outerRecord)
	directoryOffset := out.Len()
	entries := []struct {
		name   string
		data   []byte
		offset int
	}{{"[Content_Types].xml", ct, 0}, {"outer.bin", outer, len(ctRecord)}, {"inner.bin", inner, len(ctRecord) + 30 + len("outer.bin")}}
	for _, e := range entries {
		h := make([]byte, 46)
		le.PutUint32(h, 0x02014b50)
		le.PutUint16(h[4:], 20)
		le.PutUint16(h[6:], 20)
		le.PutUint32(h[16:], crc32.ChecksumIEEE(e.data))
		le.PutUint32(h[20:], uint32(len(e.data)))
		le.PutUint32(h[24:], uint32(len(e.data)))
		le.PutUint16(h[28:], uint16(len(e.name)))
		le.PutUint32(h[42:], uint32(e.offset))
		out.Write(h)
		out.WriteString(e.name)
	}
	end := make([]byte, 22)
	le.PutUint32(end, 0x06054b50)
	le.PutUint16(end[8:], uint16(len(entries)))
	le.PutUint16(end[10:], uint16(len(entries)))
	le.PutUint32(end[12:], uint32(out.Len()-directoryOffset))
	le.PutUint32(end[16:], uint32(directoryOffset))
	out.Write(end)
	return out.Bytes()
}
