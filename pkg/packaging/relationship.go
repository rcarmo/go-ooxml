package packaging

import (
	"encoding/xml"
	"fmt"
	"path"
	"strings"
)

// Relationship represents an OPC relationship.
type Relationship struct {
	ID         string     `xml:"Id,attr"`
	Type       string     `xml:"Type,attr"`
	Target     string     `xml:"Target,attr"`
	TargetMode TargetMode `xml:"TargetMode,attr,omitempty"`
}

// Relationships is a collection of relationships.
type Relationships struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []Relationship `xml:"Relationship"`
}

// NewRelationships creates an empty relationships collection.
func NewRelationships() *Relationships {
	return &Relationships{
		Relationships: make([]Relationship, 0),
	}
}

// Add adds a new relationship and returns it.
func (r *Relationships) Add(relType, target string, targetMode TargetMode) *Relationship {
	id := fmt.Sprintf("rId%d", len(r.Relationships)+1)
	rel := Relationship{
		ID:         id,
		Type:       relType,
		Target:     target,
		TargetMode: targetMode,
	}
	r.Relationships = append(r.Relationships, rel)
	return &r.Relationships[len(r.Relationships)-1]
}

// AddWithID adds a relationship with a specific ID.
func (r *Relationships) AddWithID(id, relType, target string, targetMode TargetMode) *Relationship {
	for i := range r.Relationships {
		if r.Relationships[i].ID == id {
			r.Relationships[i].Type = relType
			r.Relationships[i].Target = target
			r.Relationships[i].TargetMode = targetMode
			return &r.Relationships[i]
		}
	}

	rel := Relationship{
		ID:         id,
		Type:       relType,
		Target:     target,
		TargetMode: targetMode,
	}
	r.Relationships = append(r.Relationships, rel)
	return &r.Relationships[len(r.Relationships)-1]
}

// ByID returns the relationship with the given ID.
func (r *Relationships) ByID(id string) *Relationship {
	for i := range r.Relationships {
		if r.Relationships[i].ID == id {
			return &r.Relationships[i]
		}
	}
	return nil
}

// ByType returns all relationships of the given type.
func (r *Relationships) ByType(relType string) []*Relationship {
	var result []*Relationship
	for i := range r.Relationships {
		if r.Relationships[i].Type == relType {
			result = append(result, &r.Relationships[i])
		}
	}
	return result
}

// FirstByType returns the first relationship of the given type.
func (r *Relationships) FirstByType(relType string) *Relationship {
	for i := range r.Relationships {
		if r.Relationships[i].Type == relType {
			return &r.Relationships[i]
		}
	}
	return nil
}

// Remove removes a relationship by ID.
func (r *Relationships) Remove(id string) bool {
	for i := range r.Relationships {
		if r.Relationships[i].ID == id {
			r.Relationships = append(r.Relationships[:i], r.Relationships[i+1:]...)
			return true
		}
	}
	return false
}

// NextID returns the next available relationship ID.
func (r *Relationships) NextID() string {
	maxID := 0
	for _, rel := range r.Relationships {
		if strings.HasPrefix(rel.ID, "rId") {
			var num int
			fmt.Sscanf(rel.ID, "rId%d", &num)
			if num > maxID {
				maxID = num
			}
		}
	}
	return fmt.Sprintf("rId%d", maxID+1)
}

// MarshalXML implements custom XML marshaling for TargetMode.
func (tm TargetMode) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if tm == TargetModeExternal {
		return xml.Attr{Name: name, Value: "External"}, nil
	}
	return xml.Attr{}, nil
}

// UnmarshalXMLAttr implements custom XML unmarshaling for TargetMode.
func (tm *TargetMode) UnmarshalXMLAttr(attr xml.Attr) error {
	if attr.Value == "External" {
		*tm = TargetModeExternal
	} else {
		*tm = TargetModeInternal
	}
	return nil
}

// RelationshipsPathForPart returns the .rels file path for a given part URI.
func RelationshipsPathForPart(partURI string) string {
	dir := path.Dir(partURI)
	base := path.Base(partURI)
	if dir == "." || dir == "/" {
		return "_rels/" + base + ".rels"
	}
	return dir + "/_rels/" + base + ".rels"
}

// ResolveRelationshipTarget resolves a relationship target relative to a source part.
func ResolveRelationshipTarget(sourcePart, target string) string {
	if strings.HasPrefix(target, "/") {
		return target[1:] // Remove leading slash
	}
	sourceDir := path.Dir(sourcePart)
	if sourceDir == "." {
		return target
	}
	return path.Clean(sourceDir + "/" + target)
}
