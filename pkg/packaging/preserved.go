package packaging

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"io"
	"sort"
	"strings"
	"time"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
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
// Replace changes existing payloads; GraphPlan can add payloads and retarget
// internal edges. Neither establishes format-level dependency semantics.
// Mutable sessions are single-owner and are not safe for concurrent mutation.
type Preserved struct {
	source      []byte
	parts       map[string][]byte
	changed     map[string][]byte
	signed      bool
	deleted     map[string]bool
	added       map[string][]byte
	generation  uint64
	graphEdited bool
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
	if p.deleted[name] {
		return nil, "", &Refusal{Kind: "missing_target", Operation: "read", Part: name}
	}
	b, ok := p.changed[name]
	if !ok {
		b, ok = p.parts[name]
	}
	if !ok {
		b, ok = p.added[name]
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
	mutated := false
	for name, data := range staged {
		current, _, _ := p.Part(name)
		if bytes.Equal(current, data) {
			continue
		}
		mutated = true
		if _, added := p.added[name]; added {
			p.added[name] = data
			continue
		}
		if bytes.Equal(data, p.parts[name]) {
			delete(p.changed, name)
		} else {
			p.changed[name] = data
		}
	}
	if mutated {
		p.generation++
	}
	return nil
}

func validateXMLPayload(data []byte) error {
	if _, err := losslessxml.Parse(data); err != nil {
		return err
	}
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
	if len(p.changed) == 0 && len(p.added) == 0 && len(p.deleted) == 0 {
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
		if p.deleted[f.Name] {
			continue
		}
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
	// Additions follow the retained member order in stable name order.
	names := make([]string, 0, len(p.added))
	for name := range p.added {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		h := &zip.FileHeader{Name: name, Method: zip.Deflate}
		h.SetModTime(time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC))
		h.SetMode(0644)
		entry, err := zw.CreateHeader(h)
		if err != nil {
			return err
		}
		if _, err = entry.Write(p.added[name]); err != nil {
			return err
		}
	}
	return zw.Close()
}

// PartChange reports a byte-level member payload change, not a semantic edit.
type PartChange struct {
	Part         string `json:"part"`
	BeforeSHA256 string `json:"before_sha256"`
	AfterSHA256  string `json:"after_sha256"`
	Operation    string `json:"operation,omitempty"`
}

// Receipt is relative to the immutable input of a retained-source session.
type Receipt struct {
	Schema  int          `json:"schema"`
	Changes []PartChange `json:"changes"`
}

// Receipt reports staged payload deltas against the original source; it does
// not imply delivery. SaveAs returns it only after verified delivery succeeds.
func (p *Preserved) Receipt() Receipt {
	r := Receipt{Schema: 1, Changes: []PartChange{}}
	for name, data := range p.changed {
		r.Changes = append(r.Changes, PartChange{Part: name, BeforeSHA256: fingerprint(p.parts[name]), AfterSHA256: fingerprint(data)})
	}
	if len(p.added) > 0 || len(p.deleted) > 0 {
		r.Schema = 2
		for i := range r.Changes {
			r.Changes[i].Operation = "replace"
		}
		for name, data := range p.added {
			r.Changes = append(r.Changes, PartChange{Part: name, BeforeSHA256: "", AfterSHA256: fingerprint(data), Operation: "add"})
		}
	}
	for name := range p.deleted {
		r.Changes = append(r.Changes, PartChange{Part: name, BeforeSHA256: fingerprint(p.parts[name]), AfterSHA256: "", Operation: "delete"})
	}
	sort.Slice(r.Changes, func(i, j int) bool { return r.Changes[i].Part < r.Changes[j].Part })
	return r
}

// SaveAs verifies the completed temporary archive before replacing a regular
// destination. Failure does not consume staged changes. The receipt remains
// relative to the session's original source across successive saves.
func (p *Preserved) SaveAs(path string) (Receipt, error) {
	verify := func(temporary string) error {
		q, err := Open(temporary)
		if err != nil {
			return err
		}
		defer q.Close()
		expectedParts := p.currentParts()
		if len(q.parts) != len(expectedParts) {
			return invalidPart("save", "", "member inventory changed")
		}
		for name, original := range expectedParts {
			expected := original
			if b, ok := p.changed[name]; ok {
				expected = b
			}
			part, ok := q.parts[name]
			if !ok || !bytes.Equal(part.content, expected) {
				return invalidPart("save", name, "delivered payload mismatch")
			}
		}
		if p.graphEdited {
			candidate := &Preserved{parts: map[string][]byte{}}
			for name, part := range q.parts {
				candidate.parts[name] = part.content
			}
			if _, err := candidate.Graph(); err != nil {
				return err
			}
		}
		return nil
	}
	if err := atomicDeliver(path, p.WriteTo, verify); err != nil {
		return Receipt{}, err
	}
	return p.Receipt(), nil
}

func fingerprint(data []byte) string { s := sha256.Sum256(data); return hex.EncodeToString(s[:]) }
