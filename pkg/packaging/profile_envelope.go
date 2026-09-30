package packaging

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// Envelope is an opt-in, byte-copied OPC profile for strict registry admission
// and guarded delivery. It does not change the legacy Package/Preserved APIs.
type Envelope struct {
	source  []byte
	entries []ZIP32Entry
	deleted map[string]bool
}

// OpenEnvelope validates both ZIP structure and the supported OPC registries.
// A failure never returns a partial model or changes caller bytes.
func OpenEnvelope(source []byte) (*Envelope, error) {
	entries, err := ReadZIP32(source, ZIP32Limits{})
	if err != nil {
		return nil, err
	}
	if err = validateEnvelope(entries, nil); err != nil {
		return nil, err
	}
	return &Envelope{source: bytes.Clone(source), entries: entries, deleted: map[string]bool{}}, nil
}

// DeletePart stages a deletion; SaveAs checks the resulting relationship graph
// before it can replace an existing file. The original input is unaffected.
func (e *Envelope) DeletePart(name string) error {
	if e == nil {
		return fmt.Errorf("nil OPC envelope")
	}
	for _, entry := range e.entries {
		if entry.Name == name && !e.deleted[name] {
			e.deleted[name] = true
			return nil
		}
	}
	return &ProfileError{Reason: "opc-relationship-target-missing", Part: name}
}

// SaveAs preflights the staged graph, writes into a temporary regular file,
// verifies the complete result, then atomically renames it over the destination.
// A symlink is refused without following or replacing its target.
func (e *Envelope) SaveAs(destination string) error {
	if e == nil {
		return fmt.Errorf("nil OPC envelope")
	}
	if destination == "" {
		return fmt.Errorf("empty OPC destination")
	}
	if err := validateEnvelope(e.entries, e.deleted); err != nil {
		return err
	}
	pathName := filepath.Clean(destination)
	if info, err := os.Lstat(pathName); err == nil {
		if info.Mode()&os.ModeSymlink != 0 {
			return &ProfileError{Reason: "opc-symlink-destination", Part: pathName}
		}
		if !info.Mode().IsRegular() {
			return fmt.Errorf("destination is not a regular file: %s", pathName)
		}
	} else if !os.IsNotExist(err) {
		return err
	}
	out := bytes.Clone(e.source)
	if len(e.deleted) > 0 {
		remaining := make([]ZIP32Entry, 0, len(e.entries))
		for _, entry := range e.entries {
			if !e.deleted[entry.Name] {
				remaining = append(remaining, entry)
			}
		}
		var err error
		out, err = WriteZIP32(remaining)
		if err != nil {
			return err
		}
	}
	return atomicDeliver(pathName, func(w io.Writer) error {
		n, err := w.Write(out)
		if err == nil && n != len(out) {
			return io.ErrShortWrite
		}
		return err
	}, func(temp string) error {
		data, err := os.ReadFile(temp)
		if err != nil {
			return err
		}
		_, err = OpenEnvelope(data)
		return err
	})
}

func validateEnvelope(entries []ZIP32Entry, deleted map[string]bool) error {
	parts := map[string][]byte{}
	for _, entry := range entries {
		if deleted[entry.Name] || strings.HasSuffix(entry.Name, "/") {
			continue
		}
		if strings.Contains(entry.Name, "%") {
			return &ProfileError{Reason: "opc-part-name-invalid", Part: entry.Name}
		}
		parts[entry.Name] = entry.Data
	}
	content, ok := parts[ContentTypesPath]
	if !ok {
		return &ProfileError{Reason: "opc-content-types-invalid", Part: ContentTypesPath}
	}
	if err := validateEnvelopeRegistry(content, ctNS, "Types", func(start xml.StartElement) error {
		switch start.Name {
		case xml.Name{Space: ctNS, Local: "Default"}, xml.Name{Space: ctNS, Local: "Override"}:
			return nil
		}
		return &ProfileError{Reason: "opc-content-types-invalid", Part: ContentTypesPath}
	}); err != nil {
		return err
	}
	defaults, overrides := map[string]bool{}, map[string]bool{}
	decoder := xml.NewDecoder(bytes.NewReader(content))
	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return &ProfileError{Reason: "opc-content-types-invalid", Part: ContentTypesPath, Cause: err}
		}
		s, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		if s.Name.Space != ctNS {
			return &ProfileError{Reason: "opc-content-types-invalid", Part: ContentTypesPath}
		}
		switch s.Name.Local {
		case "Default":
			ext, hasExt := xmlAttr(s.Attr, "Extension")
			_, hasType := xmlAttr(s.Attr, "ContentType")
			if !hasExt || !hasType || ext == "" || defaults[strings.ToLower(ext)] {
				return &ProfileError{Reason: "opc-content-types-invalid", Part: ContentTypesPath}
			}
			defaults[strings.ToLower(ext)] = true
		case "Override":
			part, hasPart := xmlAttr(s.Attr, "PartName")
			_, hasType := xmlAttr(s.Attr, "ContentType")
			if !hasPart || !hasType || !strings.HasPrefix(part, "/") || strings.Contains(part, "%") || overrides[part] {
				return &ProfileError{Reason: "opc-content-types-invalid", Part: ContentTypesPath}
			}
			overrides[part] = true
		}
	}
	for name, payload := range parts {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		if err := validateEnvelopeRegistry(payload, relNS, "Relationships", nil); err != nil {
			return err
		}
		seen := map[string]bool{}
		d := xml.NewDecoder(bytes.NewReader(payload))
		source := ""
		if name != PackageRelsPath {
			if path.Base(path.Dir(name)) != "_rels" {
				return &ProfileError{Reason: "opc-target-invalid", Part: name}
			}
			source = path.Join(path.Dir(path.Dir(name)), strings.TrimSuffix(path.Base(name), ".rels"))
		}
		for {
			tok, err := d.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				return &ProfileError{Reason: "opc-target-invalid", Part: name, Cause: err}
			}
			s, ok := tok.(xml.StartElement)
			if !ok || s.Name.Local != "Relationship" {
				continue
			}
			if s.Name.Space != relNS {
				return &ProfileError{Reason: "opc-target-invalid", Part: name}
			}
			id, _ := xmlAttr(s.Attr, "Id")
			target, _ := xmlAttr(s.Attr, "Target")
			mode, _ := xmlAttr(s.Attr, "TargetMode")
			if id == "" || seen[id] {
				return &ProfileError{Reason: "opc-relationship-duplicate", Part: name}
			}
			seen[id] = true
			if mode == "External" {
				continue
			}
			if target == "" || strings.ContainsAny(target, "%\\:?#") {
				return &ProfileError{Reason: "opc-target-invalid", Part: name}
			}
			resolved := strings.TrimPrefix(target, "/")
			if !strings.HasPrefix(target, "/") {
				resolved = path.Join(path.Dir(source), target)
			}
			if err := validateMemberName(resolved, false); err != nil {
				return &ProfileError{Reason: "opc-target-invalid", Part: name}
			}
			if _, ok := parts[resolved]; !ok {
				return &ProfileError{Reason: "opc-relationship-target-missing", Part: resolved}
			}
		}
	}
	return nil
}

func xmlAttr(attrs []xml.Attr, name string) (string, bool) {
	for _, a := range attrs {
		if a.Name.Space == "" && a.Name.Local == name {
			return a.Value, true
		}
	}
	return "", false
}

func validateEnvelopeRegistry(data []byte, namespace, root string, child func(xml.StartElement) error) error {
	if _, err := losslessxml.Parse(data); err != nil {
		return &ProfileError{Reason: "opc-content-types-invalid", Cause: err}
	}
	d := xml.NewDecoder(bytes.NewReader(data))
	level := 0
	for {
		tok, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return &ProfileError{Reason: "opc-content-types-invalid", Cause: err}
		}
		switch v := tok.(type) {
		case xml.StartElement:
			if level == 0 && v.Name != (xml.Name{Space: namespace, Local: root}) {
				return &ProfileError{Reason: "opc-content-types-invalid"}
			}
			if level == 1 && child != nil {
				if err := child(v); err != nil {
					return err
				}
			}
			level++
		case xml.EndElement:
			level--
		}
	}
	return nil
}
