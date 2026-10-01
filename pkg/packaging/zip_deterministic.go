package packaging

import (
	"archive/zip"
	"bytes"
	"io"
	"math"
	"strings"
	"time"
)

// WriteZIP32Deterministic is an explicit ZIP32 writer profile. It retains
// caller order, stores small payloads/directories, deflates larger payloads,
// and sets fixed ZIP timestamps and UTF-8 flags. It never mutates entries or
// returns a partial archive on refusal. WriteZIP32 retains legacy behavior.
func WriteZIP32Deterministic(entries []ZIP32Entry) ([]byte, error) {
	// ZIP64 is deliberately outside this profile. In particular archive/zip
	// transparently emits ZIP64 end records at the 16-bit entry sentinel.
	if len(entries) >= math.MaxUint16 {
		return nil, zipReason("zip-structure-invalid", "")
	}
	seen := map[string]bool{}
	for _, entry := range entries {
		if err := validateMemberName(entry.Name, strings.HasSuffix(entry.Name, "/")); err != nil {
			return nil, zipReason("zip-name-invalid", entry.Name)
		}
		folded := asciiFold(entry.Name)
		if seen[folded] {
			return nil, zipReason("zip-case-collision", entry.Name)
		}
		seen[folded] = true
		if strings.HasSuffix(entry.Name, "/") && len(entry.Data) != 0 {
			return nil, zipReason("zip-directory-entry-invalid", entry.Name)
		}
		if uint64(len(entry.Data)) >= math.MaxUint32 {
			return nil, zipReason("zip-structure-invalid", entry.Name)
		}
	}
	var out bytes.Buffer
	writer := zip.NewWriter(&out)
	for _, entry := range entries {
		method := uint16(zip.Store)
		if len(entry.Data) > 256 {
			method = zip.Deflate
		}
		header := &zip.FileHeader{Name: entry.Name, Method: method, Flags: 0x0800}
		header.SetModTime(time.Date(1980, 1, 1, 0, 0, 0, 0, time.UTC))
		stream, err := writer.CreateHeader(header)
		if err != nil {
			return nil, err
		}
		n, err := stream.Write(entry.Data)
		if err != nil {
			return nil, err
		}
		if n != len(entry.Data) {
			return nil, io.ErrShortWrite
		}
	}
	if err := writer.Close(); err != nil {
		return nil, err
	}
	// Bound physical output too: archive/zip can otherwise transparently emit
	// ZIP64 offsets even when every individual payload fits a 32-bit field.
	if uint64(out.Len()) >= math.MaxUint32 {
		return nil, zipReason("zip-structure-invalid", "")
	}
	// The same strict physical reader verifies offset/size records, absence of
	// ZIP64 sentinels and exact geometry before any result escapes this writer.
	if _, err := ReadZIP32(out.Bytes(), ZIP32Limits{MaxArchiveBytes: math.MaxInt64, MaxEntries: math.MaxUint16 - 1, MaxEntryBytes: math.MaxUint32 - 1, MaxTotalBytes: math.MaxUint32 - 1, MaxCompressionRatio: math.MaxFloat64}); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}
