package spreadsheet

import (
	"encoding/xml"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type validationStringTable struct {
	loaded bool
	values []string
}

func (s *EditSession) validationSharedStrings() ([]string, error) {
	g, err := s.pkg.Graph()
	if err != nil {
		return nil, err
	}
	part := ""
	types := map[string]string{}
	for _, p := range g.Parts {
		types[p.Name] = p.ContentType
	}
	for _, e := range g.Edges {
		if e.Source != s.main || e.Type != packaging.RelTypeSharedStrings {
			continue
		}
		if part != "" || e.External || strings.Contains(e.Target, "#") {
			return nil, editRefusal("unsupported_structure", "ambiguous or external shared string table")
		}
		part = e.ResolvedPart
	}
	if part == "" {
		return nil, editRefusal("missing_target", "shared string relationship absent")
	}
	if types[part] != packaging.ContentTypeSharedStrings {
		return nil, editRefusal("unsupported_structure", "shared string content type mismatch")
	}
	data, _, err := s.pkg.Part(part)
	if err != nil {
		return nil, err
	}
	d, err := losslessxml.Parse(data)
	if err != nil {
		return nil, editRefusal("unsupported_structure", err.Error())
	}
	root := d.Elements()[0]
	if root.Name() != expanded("sst") {
		return nil, editRefusal("unsupported_structure", "shared string root required")
	}
	if err = validationAttrs(root, "count uniqueCount"); err != nil {
		return nil, err
	}
	if err = validationWhitespace(root); err != nil {
		return nil, err
	}
	v := validationXML{doc: d, root: root, children: map[losslessxml.Element][]losslessxml.Element{}}
	for _, e := range d.Elements() {
		if e.Name().Space != packaging.NSSpreadsheetML {
			return nil, editRefusal("unsupported_structure", "extended shared string namespace")
		}
		if p, ok := e.Parent(); ok {
			v.children[p] = append(v.children[p], e)
		}
	}
	values := []string{}
	for _, entry := range v.children[root] {
		if entry.Name() != expanded("si") {
			return nil, editRefusal("unsupported_structure", "unknown shared string entry")
		}
		value, err := v.stringValue(entry)
		if err != nil {
			return nil, err
		}
		values = append(values, value)
	}
	if count, ok := attributeValue(root, "uniqueCount"); ok {
		n, err := unsignedIndex(count)
		if err != nil || n != uint64(len(values)) {
			return nil, editRefusal("unsupported_structure", "shared string entry count mismatch")
		}
	}
	if count, ok := attributeValue(root, "count"); ok {
		if _, err := unsignedIndex(count); err != nil {
			return nil, editRefusal("unsupported_structure", "invalid shared string reference count")
		}
	}
	return values, nil
}
func validationText(e losslessxml.Element) (string, error) {
	text, leaf := e.Text()
	if !leaf {
		return "", editRefusal("unsupported_structure", "mixed string text")
	}
	for _, a := range e.Attributes() {
		if a.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || (a.Value != "preserve" && a.Value != "default") {
			return "", editRefusal("unsupported_structure", "unknown string text attribute")
		}
	}
	return text, nil
}
func (v validationXML) stringValue(container losslessxml.Element) (string, error) {
	if len(container.Attributes()) != 0 {
		return "", editRefusal("unsupported_structure", "extended string container")
	}
	if err := validationWhitespace(container); err != nil {
		return "", err
	}
	children := v.children[container]
	if len(children) == 1 && children[0].Name() == expanded("t") {
		return validationText(children[0])
	}
	if len(children) == 0 {
		return "", nil
	}
	var text strings.Builder
	for _, run := range children {
		if run.Name() != expanded("r") || len(run.Attributes()) != 0 {
			return "", editRefusal("unsupported_structure", "mixed plain/rich or extended string")
		}
		if err := validationWhitespace(run); err != nil {
			return "", err
		}
		seenPr, seenText := false, false
		for _, child := range v.children[run] {
			switch child.Name() {
			case expanded("rPr"):
				if seenPr || seenText {
					return "", editRefusal("unsupported_structure", "duplicate or misplaced string properties")
				}
				seenPr = true
				if err := v.stringProperties(child); err != nil {
					return "", err
				}
			case expanded("t"):
				if seenText {
					return "", editRefusal("unsupported_structure", "duplicate run text")
				}
				seenText = true
				t, err := validationText(child)
				if err != nil {
					return "", err
				}
				text.WriteString(t)
			default:
				return "", editRefusal("unsupported_structure", "unsupported rich string child")
			}
		}
		if !seenText {
			return "", editRefusal("unsupported_structure", "rich run has no text")
		}
	}
	return text.String(), nil
}
func (v validationXML) stringProperties(e losslessxml.Element) error {
	if len(e.Attributes()) != 0 {
		return editRefusal("unsupported_structure", "extended rich string properties")
	}
	if err := validationWhitespace(e); err != nil {
		return err
	}
	seen := map[string]bool{}
	for _, p := range v.children[e] {
		local := p.Name().Local
		if p.Name().Space != packaging.NSSpreadsheetML || seen[local] {
			return editRefusal("unsupported_structure", "duplicate rich string property")
		}
		seen[local] = true
		allowed := "val"
		switch local {
		case "rFont", "charset", "family", "b", "i", "strike", "outline", "shadow", "condense", "extend", "sz", "u", "vertAlign", "scheme":
		case "color":
			allowed = "rgb indexed theme tint auto"
		default:
			return editRefusal("unsupported_structure", "unknown rich string property")
		}
		if err := validationAttrs(p, allowed); err != nil {
			return err
		}
		if len(v.children[p]) != 0 {
			return editRefusal("unsupported_structure", "nested rich string property")
		}
		if err := validationWhitespace(p); err != nil {
			return err
		}
	}
	return nil
}
