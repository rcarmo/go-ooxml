package packaging

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"io"
	"strings"
)

// Replacement is a guarded replacement of an existing member. ExpectedSHA256 is
// the fingerprint returned by Part for the current session payload. Empty or
// mismatched fingerprints refuse; Data is copied before the transaction commits.
type Replacement struct {
	Part           string
	ExpectedSHA256 string
	Data           []byte
}

// Preserved retains a source package independently of legacy mutable models.
// This low-level API replaces existing payloads only; it does not establish
// format-level semantic validity, resolve references or infer affected parts.
// Mutable sessions are single-owner and are not safe for concurrent mutation.
type Preserved struct {
	source  []byte
	parts   map[string][]byte
	changed map[string][]byte
	signed  bool
}

// OpenPreserved takes an immutable copy of source after enforcing intake limits.
func OpenPreserved(source []byte, limits Limits) (*Preserved, error) {
	if limits.MaxSourceBytes > 0 && int64(len(source)) > limits.MaxSourceBytes {
		return nil, limitRefusal("", "source bytes")
	}
	q, err := OpenReaderWithLimits(bytes.NewReader(source), int64(len(source)), limits)
	if err != nil {
		return nil, err
	}
	defer q.Close()
	p := &Preserved{source: bytes.Clone(source), parts: map[string][]byte{}, changed: map[string][]byte{}}
	for name, part := range q.parts {
		p.parts[name] = bytes.Clone(part.content)
		if strings.HasPrefix(strings.ToLower(name), "_xmlsignatures/") {
			p.signed = true
		}
	}
	return p, nil
}

// Part returns a private copy and its current fingerprint. No mutable slice from
// the session escapes, so held payload snapshots remain valid after refusals.
func (p *Preserved) Part(name string) ([]byte, string, error) {
	b, ok := p.changed[name]
	if !ok {
		b, ok = p.parts[name]
	}
	if !ok {
		return nil, "", &Refusal{Kind: "missing_target", Operation: "read", Part: name}
	}
	return bytes.Clone(b), fingerprint(b), nil
}

// Replace preflights every member before changing any session state. Registry
// edits, signatures and new parts require future relationship-aware operations.
// Well-formed XML is checked; OOXML schema/application semantics are not implied.
func (p *Preserved) Replace(changes []Replacement) error {
	if len(changes) == 0 {
		return nil
	}
	if p.signed {
		return &Refusal{Kind: "unsupported_structure", Operation: "replace", Detail: "editing signed packages requires an explicit signature policy"}
	}
	staged := make(map[string][]byte, len(changes))
	for _, change := range changes {
		if err := validateMemberName(change.Part, false); err != nil {
			return err
		}
		if _, ok := staged[change.Part]; ok {
			return &Refusal{Kind: "ambiguous_target", Operation: "replace", Part: change.Part, Detail: "duplicate batch target"}
		}
		if change.Part == ContentTypesPath || strings.HasSuffix(change.Part, ".rels") {
			return &Refusal{Kind: "relationship_policy", Operation: "replace", Part: change.Part, Detail: "registry replacement requires a graph operation"}
		}
		_, hash, err := p.Part(change.Part)
		if err != nil {
			return err
		}
		if change.ExpectedSHA256 == "" || change.ExpectedSHA256 != hash {
			return &Refusal{Kind: "stale_target", Operation: "replace", Part: change.Part}
		}
		data := bytes.Clone(change.Data)
		if strings.HasSuffix(strings.ToLower(change.Part), ".xml") {
			if err := validateXMLPayload(data); err != nil {
				return invalidPart("replace", change.Part, err.Error())
			}
		}
		staged[change.Part] = data
	}
	for name, data := range staged {
		if bytes.Equal(data, p.parts[name]) {
			delete(p.changed, name)
		} else {
			p.changed[name] = data
		}
	}
	return nil
}

func validateXMLPayload(data []byte) error {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, roots := 0, 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		switch v := token.(type) {
		case xml.StartElement:
			if depth == 0 {
				roots++
			}
			depth++
		case xml.EndElement:
			depth--
		case xml.Directive:
			return &Refusal{Kind: "unsupported_structure", Operation: "parse_xml", Detail: "XML directives not allowed"}
		case xml.CharData:
			if depth == 0 && len(bytes.TrimSpace(v)) != 0 {
				return invalidPart("parse_xml", "", "text outside XML root")
			}
		}
	}
	if roots != 1 || depth != 0 {
		return invalidPart("parse_xml", "", "expected exactly one complete XML root")
	}
	return nil
}

// WriteTo copies the exact source for a no-op. For edits it preserves ZIP member
// order, archive comment and raw compressed data/header metadata of untouched
// members; only changed payloads are recompressed. Whole-archive identity after
// an edit is not guaranteed. Stream failures may leave partial caller output.
func (p *Preserved) WriteTo(w io.Writer) error {
	if len(p.changed) == 0 {
		n, err := w.Write(p.source)
		if err == nil && n != len(p.source) {
			return io.ErrShortWrite
		}
		return err
	}
	zr, err := zip.NewReader(bytes.NewReader(p.source), int64(len(p.source)))
	if err != nil {
		return err
	}
	zw := zip.NewWriter(w)
	if err = zw.SetComment(zr.Comment); err != nil {
		return err
	}
	for _, f := range zr.File {
		data, changed := p.changed[f.Name]
		if !changed {
			if err = zw.Copy(f); err != nil {
				return err
			}
			continue
		}
		h := f.FileHeader
		// Let CreateHeader recompute compressed size/CRC and descriptor flags.
		h.CRC32 = 0
		h.CompressedSize = 0
		h.CompressedSize64 = 0
		h.UncompressedSize = 0
		h.UncompressedSize64 = 0
		// ZIP64 size extras belong to the old payload. Other extras are retained.
		var extra []byte
		for tail := h.Extra; len(tail) >= 4; {
			n := int(little.Uint16(tail[2:4]))
			if n+4 > len(tail) {
				return invalidPart("save", f.Name, "malformed ZIP extra")
			}
			if little.Uint16(tail[:2]) != 1 {
				extra = append(extra, tail[:n+4]...)
			}
			tail = tail[n+4:]
		}
		h.Extra = extra
		entry, err := zw.CreateHeader(&h)
		if err != nil {
			return err
		}
		if _, err = entry.Write(data); err != nil {
			return err
		}
	}
	return zw.Close()
}
func fingerprint(data []byte) string { s := sha256.Sum256(data); return hex.EncodeToString(s[:]) }
