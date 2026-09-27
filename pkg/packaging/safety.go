package packaging

import (
	"fmt"
	"path"
	"strings"
	"unicode"
)

// Refusal identifies an operation rejected before delivery. It is separate from
// programmer and I/O errors and can be inspected using errors.As.
type Refusal struct {
	Kind      string
	Operation string
	Part      string
	Target    string
	Options   []string
	Detail    string
}

func (r *Refusal) Error() string {
	return fmt.Sprintf("%s: %s (%s): %s", r.Operation, r.Kind, r.Part, r.Detail)
}
func invalidPart(operation, name, detail string) error {
	return &Refusal{Kind: "invalid_package", Operation: operation, Part: name, Detail: detail}
}

// validateMemberName deliberately does not normalise aliases: accepting a
// spelling and silently changing its identity can redirect relationships.
func validateMemberName(name string, directory bool) error {
	original := name
	if directory {
		name = strings.TrimSuffix(name, "/")
	}
	if name == "" || name == "." || strings.HasPrefix(name, "/") || strings.ContainsAny(name, `\:?#`) || path.Clean(name) != name || name == ".." || strings.HasPrefix(name, "../") {
		return invalidPart("validate", original, "unsafe or non-canonical ZIP member name")
	}
	for _, r := range name {
		if unicode.IsControl(r) {
			return invalidPart("validate", original, "control character in member name")
		}
	}
	return nil
}

func (p *Package) validateOutputNames() error {
	seen := map[string]string{}
	check := func(name string) error {
		if err := validateMemberName(name, false); err != nil {
			return err
		}
		fold := strings.ToLower(name)
		if prior, ok := seen[fold]; ok {
			return invalidPart("save", name, "duplicate or case-colliding member with "+prior)
		}
		seen[fold] = name
		return nil
	}
	if err := check(ContentTypesPath); err != nil {
		return err
	}
	for source, rels := range p.relationships {
		if len(rels.Relationships) == 0 {
			continue
		}
		name := PackageRelsPath
		if source != "" && source != "." {
			name = RelationshipsPathForPart(source)
		}
		if err := check(name); err != nil {
			return err
		}
	}
	for name := range p.parts {
		if name == ContentTypesPath || strings.HasSuffix(name, ".rels") {
			continue
		}
		if err := check(name); err != nil {
			return err
		}
	}
	return nil
}
