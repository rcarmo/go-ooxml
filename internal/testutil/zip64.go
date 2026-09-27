package testutil

import (
	"bytes"
	"compress/flate"
	"encoding/binary"
	"hash/crc32"
)

// ZIP64Fixture builds a tiny deterministic OPC archive independent of archive/zip's
// writer thresholds. All entries use local and central ZIP64 size declarations.
// Descriptor is "", "signed", or "unsigned"; Deflate includes an empty member.
// Defect names intentionally alter one field/region for hostile intake tests.
type ZIP64Fixture struct {
	Descriptor string
	Deflate    bool
	Defect     string
}

func TinyZIP64(opt ZIP64Fixture) []byte {
	le := binary.LittleEndian
	var out, central bytes.Buffer
	members := []struct {
		name string
		data []byte
	}{{"[Content_Types].xml", []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/></Types>`)}, {"data.bin", []byte("zip64 payload")}, {"empty.bin", nil}}
	for index, m := range members {
		payload := m.data
		method := uint16(0)
		if opt.Deflate {
			method = 8
			var b bytes.Buffer
			w, _ := flate.NewWriter(&b, flate.DefaultCompression)
			_, _ = w.Write(payload)
			_ = w.Close()
			payload = b.Bytes()
		}
		flags := uint16(0)
		if opt.Descriptor != "" {
			flags = 8
		}
		offset := uint64(out.Len())
		crc := crc32.ChecksumIEEE(m.data)
		extra := make([]byte, 20)
		le.PutUint16(extra, 1)
		le.PutUint16(extra[2:], 16)
		le.PutUint64(extra[4:], uint64(len(m.data)))
		le.PutUint64(extra[12:], uint64(len(payload)))
		if index == 1 {
			switch opt.Defect {
			case "short local extra":
				le.PutUint16(extra[2:], 8)
				extra = extra[:12]
			case "local size mismatch":
				le.PutUint64(extra[4:], uint64(len(m.data)+1))
			case "missing local extra":
				extra = nil
			case "duplicate local extra":
				extra = append(extra, extra...)
			case "trailing local extra byte":
				extra = append(extra, 0)
			case "overflow local size":
				le.PutUint64(extra[4:], ^uint64(0))
			}
		}
		local := make([]byte, 30)
		le.PutUint32(local, 0x04034b50)
		le.PutUint16(local[4:], 45)
		le.PutUint16(local[6:], flags)
		le.PutUint16(local[8:], method)
		if flags == 0 {
			le.PutUint32(local[14:], crc)
		}
		le.PutUint32(local[18:], 0xffffffff)
		le.PutUint32(local[22:], 0xffffffff)
		le.PutUint16(local[26:], uint16(len(m.name)))
		le.PutUint16(local[28:], uint16(len(extra)))
		out.Write(local)
		out.WriteString(m.name)
		out.Write(extra)
		out.Write(payload)
		if flags != 0 {
			if opt.Descriptor == "signed" {
				_ = binary.Write(&out, le, uint32(0x08074b50))
			}
			descriptor := make([]byte, 20)
			le.PutUint32(descriptor, crc)
			le.PutUint64(descriptor[4:], uint64(len(payload)))
			le.PutUint64(descriptor[12:], uint64(len(m.data)))
			if index == 1 && opt.Defect == "descriptor mismatch" {
				le.PutUint64(descriptor[12:], 999)
			}
			out.Write(descriptor)
		}
		ce := make([]byte, 28)
		le.PutUint16(ce, 1)
		le.PutUint16(ce[2:], 24)
		le.PutUint64(ce[4:], uint64(len(m.data)))
		le.PutUint64(ce[12:], uint64(len(payload)))
		le.PutUint64(ce[20:], offset)
		if index == 1 {
			switch opt.Defect {
			case "short central extra":
				le.PutUint16(ce[2:], 16)
				ce = ce[:20]
			case "duplicate central extra":
				ce = append(ce, ce...)
			case "trailing central extra byte":
				ce = append(ce, 0)
			}
		}
		ch := make([]byte, 46)
		le.PutUint32(ch, 0x02014b50)
		le.PutUint16(ch[4:], 45)
		le.PutUint16(ch[6:], 45)
		le.PutUint16(ch[8:], flags)
		le.PutUint16(ch[10:], method)
		le.PutUint32(ch[16:], crc)
		le.PutUint32(ch[20:], 0xffffffff)
		le.PutUint32(ch[24:], 0xffffffff)
		le.PutUint16(ch[28:], uint16(len(m.name)))
		le.PutUint16(ch[30:], uint16(len(ce)))
		le.PutUint32(ch[42:], 0xffffffff)
		central.Write(ch)
		central.WriteString(m.name)
		central.Write(ce)
	}
	directoryOffset := out.Len()
	out.Write(central.Bytes())
	if opt.Defect == "gap before ZIP64 end" {
		out.WriteString("UNDECLARED")
	}
	zoff := out.Len()
	zend := make([]byte, 56)
	le.PutUint32(zend, 0x06064b50)
	le.PutUint64(zend[4:], 44)
	le.PutUint16(zend[12:], 45)
	le.PutUint16(zend[14:], 45)
	le.PutUint64(zend[24:], uint64(len(members)))
	le.PutUint64(zend[32:], uint64(len(members)))
	le.PutUint64(zend[40:], uint64(central.Len()))
	le.PutUint64(zend[48:], uint64(directoryOffset))
	switch opt.Defect {
	case "ZIP64 length overflow":
		le.PutUint64(zend[4:], ^uint64(0))
	case "ZIP64 short length":
		le.PutUint64(zend[4:], 43)
	case "ZIP64 length crosses locator":
		le.PutUint64(zend[4:], 45)
	}
	out.Write(zend)
	if opt.Defect == "gap before locator" {
		out.WriteString("GAP")
	}
	locator := make([]byte, 20)
	le.PutUint32(locator, 0x07064b50)
	le.PutUint64(locator[8:], uint64(zoff))
	le.PutUint32(locator[16:], 1)
	out.Write(locator)
	end := make([]byte, 22)
	le.PutUint32(end, 0x06054b50)
	le.PutUint16(end[8:], 0xffff)
	le.PutUint16(end[10:], 0xffff)
	le.PutUint32(end[12:], 0xffffffff)
	le.PutUint32(end[16:], 0xffffffff)
	if opt.Defect == "classic count mismatch" {
		le.PutUint16(end[8:], 1)
	}
	out.Write(end)
	return out.Bytes()
}
