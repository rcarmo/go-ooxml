package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"strings"
)

// ProfileError is a structured refusal for the opt-in ZIP32/OPC profile.
// Reason is stable; diagnostic text and native error wording are not classifiers.
type ProfileError struct {
	Reason string
	Part   string
	Cause  error
}

func (e *ProfileError) Error() string {
	if e.Cause != nil {
		return fmt.Sprintf("%s (%s): %v", e.Reason, e.Part, e.Cause)
	}
	return fmt.Sprintf("%s (%s)", e.Reason, e.Part)
}
func (e *ProfileError) Unwrap() error     { return e.Cause }
func zipReason(reason, part string) error { return &ProfileError{Reason: reason, Part: part} }

// asciiFold applies exactly the ASCII member-name collision rule.
func asciiFold(name string) string {
	b := []byte(name)
	for i, ch := range b {
		if ch >= 'A' && ch <= 'Z' {
			b[i] = ch + ('a' - 'A')
		}
	}
	return string(b)
}

// ZIP32Entry is an ordered, detached member payload. Directories carry no data.
type ZIP32Entry struct {
	Name string
	Data []byte
}

// ZIP32Limits bounds the opt-in raw reader. Zero selects a finite default;
// explicit positive values override only the selected limit.
type ZIP32Limits struct {
	MaxArchiveBytes     int64
	MaxEntries          int
	MaxEntryBytes       uint64
	MaxTotalBytes       uint64
	MaxCompressionRatio float64
}

// Conservative finite defaults apply even to a zero-value ZIP32Limits.
const (
	defaultZIP32ArchiveBytes     int64  = 64 << 20
	defaultZIP32Entries                 = 4096
	defaultZIP32EntryBytes       uint64 = 16 << 20
	defaultZIP32TotalBytes       uint64 = 64 << 20
	defaultZIP32CompressionRatio        = 1000.0
)

func boundedZIP32Limits(limits ZIP32Limits) (ZIP32Limits, error) {
	if limits.MaxArchiveBytes < 0 || limits.MaxEntries < 0 || math.IsNaN(limits.MaxCompressionRatio) || math.IsInf(limits.MaxCompressionRatio, 0) || limits.MaxCompressionRatio < 0 {
		return ZIP32Limits{}, fmt.Errorf("invalid ZIP32 budget")
	}
	if limits.MaxArchiveBytes == 0 {
		limits.MaxArchiveBytes = defaultZIP32ArchiveBytes
	}
	if limits.MaxEntries == 0 {
		limits.MaxEntries = defaultZIP32Entries
	}
	if limits.MaxEntryBytes == 0 {
		limits.MaxEntryBytes = defaultZIP32EntryBytes
	}
	if limits.MaxTotalBytes == 0 {
		limits.MaxTotalBytes = defaultZIP32TotalBytes
	}
	if limits.MaxCompressionRatio == 0 {
		limits.MaxCompressionRatio = defaultZIP32CompressionRatio
	}
	return limits, nil
}

// ReadZIP32 reads a standalone ZIP32 archive without requiring an OPC content
// type registry. It preflights central/local records before letting archive/zip
// expand payloads, and never returns a partial result on refusal.
func ReadZIP32(source []byte, limits ZIP32Limits) ([]ZIP32Entry, error) {
	var err error
	limits, err = boundedZIP32Limits(limits)
	if err != nil {
		return nil, err
	}
	n := len(source)
	if limits.MaxArchiveBytes > 0 && int64(n) > limits.MaxArchiveBytes {
		return nil, zipReason("zip-archive-too-large", "")
	}
	if n < 22 {
		return nil, zipReason("zip-end-record-missing", "")
	}
	// A declared comment must consume the complete suffix, never conceal bytes.
	end := -1
	for at := n - 22; at >= 0 && at >= n-22-65535; at-- {
		if binary.LittleEndian.Uint32(source[at:at+4]) == 0x06054b50 && at+22+int(binary.LittleEndian.Uint16(source[at+20:at+22])) == n {
			end = at
			break
		}
	}
	if end < 0 {
		return nil, zipReason("zip-end-record-missing", "")
	}
	e := source[end:]
	u16 := binary.LittleEndian.Uint16
	u32 := binary.LittleEndian.Uint32
	if u16(e[4:6]) != 0 || u16(e[6:8]) != 0 || u16(e[8:10]) != u16(e[10:12]) {
		return nil, zipReason("zip-multi-disk-unsupported", "")
	}
	count := int(u16(e[10:12]))
	if count == math.MaxUint16 || u32(e[12:16]) == math.MaxUint32 || u32(e[16:20]) == math.MaxUint32 {
		return nil, zipReason("zip-structure-invalid", "")
	}
	if limits.MaxEntries > 0 && count > limits.MaxEntries {
		return nil, zipReason("zip-too-many-entries", "")
	}
	start, length := int(u32(e[16:20])), int(u32(e[12:16]))
	if start > end || length != end-start {
		return nil, zipReason("zip-structure-invalid", "")
	}
	cursor := start
	seen := map[string]bool{}
	folded := map[string]bool{}
	entries := make([]ZIP32Entry, 0, count)
	var total uint64
	for i := 0; i < count; i++ {
		if cursor > end || end-cursor < 46 || u32(source[cursor:cursor+4]) != 0x02014b50 {
			return nil, zipReason("zip-structure-invalid", "")
		}
		h := source[cursor : cursor+46]
		nameLen, extraLen, commentLen := int(u16(h[28:30])), int(u16(h[30:32])), int(u16(h[32:34]))
		if nameLen+extraLen+commentLen > end-cursor-46 {
			return nil, zipReason("zip-structure-invalid", "")
		}
		name := string(source[cursor+46 : cursor+46+nameLen])
		cursor += 46 + nameLen + extraLen + commentLen
		if seen[name] {
			return nil, zipReason("zip-duplicate-entry", name)
		}
		if folded[asciiFold(name)] {
			return nil, zipReason("zip-case-collision", name)
		}
		seen[name], folded[asciiFold(name)] = true, true
		directory := strings.HasSuffix(name, "/")
		if err := validateMemberName(name, directory); err != nil {
			return nil, zipReason("zip-name-invalid", name)
		}
		flags, method := u16(h[8:10]), u16(h[10:12])
		if flags&1 != 0 {
			return nil, zipReason("zip-encryption-unsupported", name)
		}
		if method != zip.Store && method != zip.Deflate {
			return nil, zipReason("zip-method-unsupported", name)
		}
		if u16(h[34:36]) != 0 {
			return nil, zipReason("zip-multi-disk-unsupported", name)
		}
		compressed, expanded := uint64(u32(h[20:24])), uint64(u32(h[24:28]))
		if compressed == math.MaxUint32 || expanded == math.MaxUint32 || u32(h[42:46]) == math.MaxUint32 {
			return nil, zipReason("zip-structure-invalid", name)
		}
		if directory && (compressed != 0 || expanded != 0) {
			return nil, zipReason("zip-directory-entry-invalid", name)
		}
		if !directory {
			if limits.MaxEntryBytes > 0 && expanded > limits.MaxEntryBytes {
				return nil, zipReason("zip-entry-too-large", name)
			}
			if math.MaxUint64-total < expanded {
				return nil, zipReason("zip-total-too-large", name)
			}
			total += expanded
			if limits.MaxTotalBytes > 0 && total > limits.MaxTotalBytes {
				return nil, zipReason("zip-total-too-large", name)
			}
			if limits.MaxCompressionRatio > 0 && (compressed == 0 && expanded > 0 || compressed > 0 && float64(expanded)/float64(compressed) > limits.MaxCompressionRatio) {
				return nil, zipReason("zip-compression-ratio-exceeded", name)
			}
		}
		offset := int(u32(h[42:46]))
		if offset > start || start-offset < 30 || u32(source[offset:offset+4]) != 0x04034b50 {
			return nil, zipReason("zip-structure-invalid", name)
		}
		local := source[offset : offset+30]
		localNameLen, localExtraLen := int(u16(local[26:28])), int(u16(local[28:30]))
		if localNameLen+localExtraLen > start-offset-30 {
			return nil, zipReason("zip-local-metadata-mismatch", name)
		}
		if string(source[offset+30:offset+30+localNameLen]) != name || u16(local[6:8]) != flags || u16(local[8:10]) != method {
			return nil, zipReason("zip-local-metadata-mismatch", name)
		}
		dataStart := offset + 30 + localNameLen + localExtraLen
		if compressed > uint64(start-dataStart) {
			return nil, zipReason("zip-size-mismatch", name)
		}
		if flags&8 == 0 && (u32(local[14:18]) != u32(h[16:20]) || u32(local[18:22]) != uint32(compressed) || u32(local[22:26]) != uint32(expanded)) {
			return nil, zipReason("zip-local-metadata-mismatch", name)
		}
		entries = append(entries, ZIP32Entry{Name: name})
	}
	if cursor != end {
		return nil, zipReason("zip-structure-invalid", "")
	}
	zr, err := zip.NewReader(bytes.NewReader(source), int64(n))
	if err != nil || len(zr.File) != count {
		return nil, zipReason("zip-structure-invalid", "")
	}
	if err := validateZIPStructure(bytes.NewReader(source), int64(n), zr.File); err != nil {
		return nil, zipReason("zip-structure-invalid", "")
	}
	lenPayloads := 0
	for i, f := range zr.File {
		if f.Name != entries[i].Name {
			return nil, zipReason("zip-structure-invalid", f.Name)
		}
		if f.FileInfo().IsDir() {
			continue
		}
		r, err := f.Open()
		if err != nil {
			return nil, zipReason("zip-size-mismatch", f.Name)
		}
		// The per-entry and aggregate budgets are already preflighted from
		// the directory. Bound real inflation too, so a false declared size
		// cannot allocate past either remaining budget before refusal.
		remaining := limits.MaxTotalBytes - uint64(lenPayloads)
		maximum := f.UncompressedSize64
		if maximum > limits.MaxEntryBytes {
			maximum = limits.MaxEntryBytes
		}
		if maximum > remaining {
			maximum = remaining
		}
		data, readErr := io.ReadAll(io.LimitReader(r, int64(maximum)+1))
		closeErr := r.Close()
		if readErr != nil || closeErr != nil {
			if readErr == zip.ErrChecksum {
				return nil, zipReason("zip-crc-mismatch", f.Name)
			}
			return nil, zipReason("zip-size-mismatch", f.Name)
		}
		if uint64(len(data)) != f.UncompressedSize64 {
			return nil, zipReason("zip-size-mismatch", f.Name)
		}
		lenPayloads += len(data)
		entries[i].Data = data
	}
	return entries, nil
}

// WriteZIP32 emits ordered, detached ZIP entries. It never returns partial
// archive bytes and rejects aliases before touching an output buffer.
func WriteZIP32(entries []ZIP32Entry) ([]byte, error) {
	seen := map[string]bool{}
	for _, e := range entries {
		if err := validateMemberName(e.Name, strings.HasSuffix(e.Name, "/")); err != nil {
			return nil, zipReason("zip-name-invalid", e.Name)
		}
		fold := asciiFold(e.Name)
		if seen[fold] {
			return nil, zipReason("zip-case-collision", e.Name)
		}
		seen[fold] = true
		if strings.HasSuffix(e.Name, "/") && len(e.Data) > 0 {
			return nil, zipReason("zip-directory-entry-invalid", e.Name)
		}
	}
	var b bytes.Buffer
	zw := zip.NewWriter(&b)
	for _, e := range entries {
		method := uint16(zip.Deflate)
		if strings.HasSuffix(e.Name, "/") {
			method = zip.Store
		}
		h := &zip.FileHeader{Name: e.Name, Method: method, Flags: 0x0800}
		w, err := zw.CreateHeader(h)
		if err != nil {
			return nil, err
		}
		if _, err = w.Write(e.Data); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return b.Bytes(), nil
}
