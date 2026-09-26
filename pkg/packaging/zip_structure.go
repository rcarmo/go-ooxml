package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"sort"
)

var little = binary.LittleEndian

type zipRange struct {
	start, end int64
	name       string
}

// validateZIPStructure checks the parts archive/zip intentionally leaves to
// callers: local-vs-central identity and overlapping physical member ranges.
// OOXML is a standalone ZIP, not an executable with a prepended stub.
func validateZIPStructure(r io.ReaderAt, size int64, files []*zip.File) error {
	fail := func(name, detail string) error { return invalidPart("open", name, detail) }
	read := func(offset, length int64) ([]byte, error) {
		if offset < 0 || length < 0 || offset > size || length > size-offset {
			return nil, fmt.Errorf("ZIP range outside archive")
		}
		b := make([]byte, int(length))
		_, err := r.ReadAt(b, offset)
		return b, err
	}
	n := int64(65535 + 22)
	if n > size {
		n = size
	}
	tail, err := read(size-n, n)
	if err != nil {
		return err
	}
	end := -1
	for i := len(tail) - 22; i >= 0; i-- {
		if little.Uint32(tail[i:i+4]) == 0x06054b50 && i+22+int(little.Uint16(tail[i+20:i+22])) == len(tail) {
			end = i
			break
		}
	}
	if end < 0 {
		return fail("", "missing end-of-directory record")
	}
	e := tail[end:]
	eoff := size - n + int64(end)
	if little.Uint16(e[4:6]) != 0 || little.Uint16(e[6:8]) != 0 {
		return fail("", "multi-disk archive unsupported")
	}
	count := uint64(little.Uint16(e[10:12]))
	directorySize := uint64(little.Uint32(e[12:16]))
	directoryOffset := uint64(little.Uint32(e[16:20]))
	boundary := eoff
	if count == math.MaxUint16 || directorySize == math.MaxUint32 || directoryOffset == math.MaxUint32 {
		locator, err := read(eoff-20, 20)
		if err != nil {
			return fail("", "missing ZIP64 locator")
		}
		if little.Uint32(locator[:4]) != 0x07064b50 || little.Uint32(locator[4:8]) != 0 || little.Uint32(locator[16:20]) != 1 {
			return fail("", "invalid ZIP64 locator")
		}
		at := little.Uint64(locator[8:16])
		if at > uint64(size) {
			return fail("", "ZIP64 offset outside archive")
		}
		z, err := read(int64(at), 56)
		if err != nil {
			return fail("", "truncated ZIP64 directory")
		}
		if little.Uint32(z[:4]) != 0x06064b50 || little.Uint64(z[4:12]) < 44 || little.Uint32(z[16:20]) != 0 || little.Uint32(z[20:24]) != 0 {
			return fail("", "invalid ZIP64 directory")
		}
		if little.Uint64(z[24:32]) != little.Uint64(z[32:40]) {
			return fail("", "split ZIP64 directory")
		}
		count = little.Uint64(z[32:40])
		directorySize = little.Uint64(z[40:48])
		directoryOffset = little.Uint64(z[48:56])
		boundary = int64(at)
	} else if little.Uint16(e[8:10]) != little.Uint16(e[10:12]) {
		return fail("", "split directory")
	}
	if count != uint64(len(files)) || directoryOffset > uint64(boundary) || directorySize > uint64(boundary)-directoryOffset {
		return fail("", "inconsistent directory bounds")
	}
	cursor := int64(directoryOffset)
	limit := cursor + int64(directorySize)
	ranges := make([]zipRange, 0, len(files))
	for _, f := range files {
		h, err := read(cursor, 46)
		if err != nil {
			return fail(f.Name, "truncated central header")
		}
		if little.Uint32(h[:4]) != 0x02014b50 {
			return fail(f.Name, "invalid central signature")
		}
		nameLen, extraLen, commentLen := int64(little.Uint16(h[28:30])), int64(little.Uint16(h[30:32])), int64(little.Uint16(h[32:34]))
		extent := 46 + nameLen + extraLen + commentLen
		if extent > limit-cursor {
			return fail(f.Name, "central header crosses directory boundary")
		}
		variable, err := read(cursor+46, nameLen+extraLen)
		if err != nil {
			return fail(f.Name, "truncated central metadata")
		}
		if string(variable[:nameLen]) != f.Name {
			return fail(f.Name, "central member identity mismatch")
		}
		if little.Uint16(h[34:36]) != 0 {
			return fail(f.Name, "multi-disk member unsupported")
		}
		offset := uint64(little.Uint32(h[42:46]))
		if offset == math.MaxUint32 {
			extra := variable[nameLen:]
			found := false
			for len(extra) >= 4 {
				tag, n := little.Uint16(extra[:2]), int(little.Uint16(extra[2:4]))
				extra = extra[4:]
				if n > len(extra) {
					return fail(f.Name, "truncated ZIP64 extra")
				}
				if tag == 1 {
					data := extra[:n]
					skip := 0
					if little.Uint32(h[24:28]) == math.MaxUint32 {
						skip += 8
					}
					if little.Uint32(h[20:24]) == math.MaxUint32 {
						skip += 8
					}
					if len(data) < skip+8 {
						return fail(f.Name, "missing ZIP64 local offset")
					}
					offset = little.Uint64(data[skip : skip+8])
					found = true
					break
				}
				extra = extra[n:]
			}
			if !found {
				return fail(f.Name, "missing ZIP64 local offset")
			}
		}
		if offset > directoryOffset {
			return fail(f.Name, "local header inside directory")
		}
		local, err := read(int64(offset), 30)
		if err != nil {
			return fail(f.Name, "truncated local header")
		}
		if little.Uint32(local[:4]) != 0x04034b50 {
			return fail(f.Name, "invalid local signature")
		}
		if little.Uint16(local[6:8]) != f.Flags || little.Uint16(local[8:10]) != f.Method {
			return fail(f.Name, "local flags/compression disagree with directory")
		}
		localNameLen, localExtraLen := int64(little.Uint16(local[26:28])), int64(little.Uint16(local[28:30]))
		name, err := read(int64(offset)+30, localNameLen)
		if err != nil || !bytes.Equal(name, []byte(f.Name)) {
			return fail(f.Name, "local name disagrees with directory")
		}
		dataStart := int64(offset) + 30 + localNameLen + localExtraLen
		actualOffset, err := f.DataOffset()
		if err != nil || dataStart != actualOffset {
			return fail(f.Name, "ambiguous payload offset")
		}
		if dataStart > int64(directoryOffset) || f.CompressedSize64 > uint64(int64(directoryOffset)-dataStart) {
			return fail(f.Name, "member overlaps central directory")
		}
		end := dataStart + int64(f.CompressedSize64)
		if f.Flags&8 == 0 {
			if little.Uint32(local[14:18]) != f.CRC32 {
				return fail(f.Name, "local CRC disagrees with directory")
			}
			compressed, uncompressed := uint64(little.Uint32(local[18:22])), uint64(little.Uint32(local[22:26]))
			if compressed != math.MaxUint32 && compressed != f.CompressedSize64 {
				return fail(f.Name, "local compressed size disagrees")
			}
			if uncompressed != math.MaxUint32 && uncompressed != f.UncompressedSize64 {
				return fail(f.Name, "local uncompressed size disagrees")
			}
			// ZIP64 sentinels need the local extra fields; archive/zip validates the
			// central sizes and stream. Reject ambiguous local declarations for now.
			if compressed == math.MaxUint32 || uncompressed == math.MaxUint32 {
				return fail(f.Name, "local ZIP64 sizes require unsupported verification")
			}
		} else {
			width := int64(12)
			if f.CompressedSize64 >= math.MaxUint32 || f.UncompressedSize64 >= math.MaxUint32 {
				width = 20
			}
			// CRC may itself equal the optional signature. Validate both layouts
			// against the complete directory tuple rather than guessing from CRC.
			ends := []int64{}
			for _, skip := range []int64{0, 4} {
				if skip == 4 {
					signature, err := read(end, 4)
					if err != nil || little.Uint32(signature) != 0x08074b50 {
						continue
					}
				}
				if end+skip+width > int64(directoryOffset) {
					continue
				}
				descriptor, err := read(end+skip, width)
				if err != nil || little.Uint32(descriptor[:4]) != f.CRC32 {
					continue
				}
				var c, u uint64
				if width == 12 {
					c = uint64(little.Uint32(descriptor[4:8]))
					u = uint64(little.Uint32(descriptor[8:12]))
				} else {
					c = little.Uint64(descriptor[4:12])
					u = little.Uint64(descriptor[12:20])
				}
				if c == f.CompressedSize64 && u == f.UncompressedSize64 {
					ends = append(ends, end+skip+width)
				}
			}
			if len(ends) != 1 {
				return fail(f.Name, "missing, inconsistent or ambiguous data descriptor")
			}
			end = ends[0]
		}
		if end > int64(directoryOffset) {
			return fail(f.Name, "descriptor overlaps central directory")
		}
		ranges = append(ranges, zipRange{int64(offset), end, f.Name})
		cursor += extent
	}
	if cursor != limit {
		return fail("", "unexpected directory bytes")
	}
	sort.Slice(ranges, func(i, j int) bool { return ranges[i].start < ranges[j].start })
	if len(ranges) > 0 && ranges[0].start != 0 {
		return fail(ranges[0].name, "prepended bytes are not a standalone OPC archive")
	}
	for i := 1; i < len(ranges); i++ {
		if ranges[i].start < ranges[i-1].end {
			return fail(ranges[i].name, "overlapping archive entries")
		}
	}
	return nil
}
