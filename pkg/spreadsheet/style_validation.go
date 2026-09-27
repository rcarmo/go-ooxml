package spreadsheet

import (
	"encoding/xml"
	"strconv"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// ValidateStyles checks that cell/row/column style indices resolve. It does not
// calculate effective formatting or certify all dependent font/fill/XF semantics.
// When no styles relationship exists, only implicit/default zero is accepted.
// When a table exists it must actually contain style zero, even for omitted s.
func (s *EditSession) ValidateStyles() error {
	graph, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	styles := ""
	types := map[string]string{}
	for _, p := range graph.Parts {
		types[p.Name] = p.ContentType
	}
	for _, edge := range graph.Edges {
		if edge.Source == s.main && edge.Type == packaging.RelTypeStyles {
			if edge.External || styles != "" {
				return editRefusal("unsupported_structure", "ambiguous or external styles relationship")
			}
			styles = edge.ResolvedPart
		}
	}
	count := uint64(1) // implicit default only when there is no style part
	if styles != "" {
		if types[styles] != packaging.ContentTypeExcelStyles {
			return editRefusal("unsupported_structure", "style part has wrong content type")
		}
		data, _, err := s.pkg.Part(styles)
		if err != nil {
			return err
		}
		d, err := losslessxml.Parse(data)
		if err != nil {
			return editRefusal("unsupported_structure", "invalid styles XML: "+err.Error())
		}
		elements := d.Elements()
		root := elements[0]
		if root.Name() != expanded("styleSheet") {
			return editRefusal("unsupported_structure", "styles root is not styleSheet")
		}
		var table losslessxml.Element
		tables := 0
		for _, e := range elements {
			parent, ok := e.Parent()
			if ok && parent == root && e.Name() == expanded("cellXfs") {
				table = e
				tables++
			}
		}
		if tables != 1 {
			return editRefusal("unsupported_structure", "styles requires exactly one cellXfs table")
		}
		count = 0
		for _, e := range elements {
			parent, ok := e.Parent()
			if !ok || parent != table {
				continue
			}
			if e.Name() != expanded("xf") {
				return editRefusal("unsupported_structure", "unknown cellXfs child")
			}
			count++
		}
		if declared, present := attributeValue(table, "count"); present {
			n, err := unsignedIndex(declared)
			if err != nil || n != count {
				return editRefusal("unsupported_structure", "cellXfs count disagrees with actual xf entries")
			}
		}
		if count == 0 {
			return editRefusal("unsupported_structure", "style zero is absent from cellXfs")
		}
	}
	for _, part := range graph.Parts {
		if part.ContentType != packaging.ContentTypeWorksheet {
			continue
		}
		b, _, err := s.pkg.Part(part.Name)
		if err != nil {
			return err
		}
		d, err := losslessxml.Parse(b)
		if err != nil {
			return editRefusal("unsupported_structure", "invalid worksheet XML: "+err.Error())
		}
		for _, e := range d.Elements() {
			if e.Name().Space != packaging.NSSpreadsheetML {
				continue
			}
			field := ""
			always := false
			switch e.Name().Local {
			case "c":
				field = "s"
				always = true
			case "row":
				field = "s"
			case "col":
				field = "style"
			default:
				continue
			}
			text, present := attributeValue(e, field)
			if !present && !always {
				continue
			}
			index := uint64(0)
			if present {
				index, err = unsignedIndex(text)
				if err != nil {
					return editRefusal("unsupported_structure", "malformed style index in "+part.Name)
				}
			}
			if index >= count {
				return editRefusal("unsupported_structure", "style index is outside cellXfs in "+part.Name)
			}
		}
	}
	return nil
}
func attributeValue(e losslessxml.Element, local string) (string, bool) {
	for _, a := range e.Attributes() {
		if a.Name == (xml.Name{Local: local}) {
			return a.Value, true
		}
	}
	return "", false
}
func unsignedIndex(text string) (uint64, error) {
	if text == "" {
		return 0, strconv.ErrSyntax
	}
	for _, r := range text {
		if r < '0' || r > '9' {
			return 0, strconv.ErrSyntax
		}
	}
	return strconv.ParseUint(text, 10, 32)
}
