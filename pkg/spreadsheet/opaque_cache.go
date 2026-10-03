package spreadsheet

import (
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"sort"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Only the reviewed two root edges and typed orphan-independent cache parts
// qualify. Any worksheet/workbook owner edge, outgoing edge, extra root edge,
// different URI/ID/target/type or missing override refuses before mutation.
func (s *EditSession) sealedOpaqueCachePreflight() error {
	g, err := s.pkg.Graph()
	if err != nil {
		return err
	}
	expected := map[string]struct{ typ, target, part, contentType string }{
		"cache1": {"urn:contract20:opaque/chart", "xl/charts/cache-boundary.xml", "xl/charts/cache-boundary.xml", packaging.ContentTypeChart},
		"cache2": {"urn:contract20:opaque/external", "xl/externalLinks/cache-boundary.xml", "xl/externalLinks/cache-boundary.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"},
	}
	workbookBytes, _, err := s.pkg.Part(s.main)
	if err != nil {
		return err
	}
	workbook, err := losslessxml.Parse(workbookBytes)
	if err != nil {
		return editRefusal("xlsx-opaque-cache-scope-unsupported", "invalid workbook XML")
	}
	worksheetIDs := map[string]string{}
	for _, sheet := range workbook.Elements() {
		if sheet.Name() != expanded("sheet") {
			continue
		}
		parent, ok := sheet.Parent()
		if !ok || parent.Name() != expanded("sheets") {
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "worksheet owner differs")
		}
		name := attr(sheet, "name")
		id := ""
		for _, a := range sheet.Attributes() {
			if a.Name.Space == packaging.NSDocumentRelationships && a.Name.Local == "id" {
				id = a.Value
			}
		}
		if name == "" || id == "" || worksheetIDs[id] != "" || s.sheets[name] == "" {
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "worksheet identity differs")
		}
		worksheetIDs[id] = s.sheets[name]
	}
	if len(worksheetIDs) != len(s.sheets) {
		return editRefusal("xlsx-opaque-cache-scope-unsupported", "worksheet count differs")
	}
	seenWorksheets := map[string]bool{}
	seen := map[string]bool{}
	workbookRootEdges := 0
	for _, e := range g.Edges {
		if e.Source == "" {
			if e.Type == packaging.RelTypeOfficeDocument {
				workbookRootEdges++
				if e.ID != "rId1" || e.Target != "xl/workbook.xml" || e.ResolvedPart != s.main || e.External {
					return editRefusal("xlsx-opaque-cache-scope-unsupported", "root workbook edge differs")
				}
			} else {
				want, ok := expected[e.ID]
				if !ok || seen[e.ID] || e.Type != want.typ || e.Target != want.target || e.ResolvedPart != want.part || e.External {
					return editRefusal("xlsx-opaque-cache-scope-unsupported", "unknown root cache relationship")
				}
				seen[e.ID] = true
			}
		}
		for _, want := range expected {
			if e.Source == want.part || e.ResolvedPart == want.part && e.Source != "" {
				return editRefusal("xlsx-opaque-cache-scope-unsupported", "semantic cache owner/outgoing edge")
			}
		}
		if e.Source == s.main {
			if e.External {
				return editRefusal("xlsx-opaque-cache-scope-unsupported", "external workbook dependency")
			}
			switch e.Type {
			case packaging.RelTypeWorksheet:
				if worksheetIDs[e.ID] == "" || worksheetIDs[e.ID] != e.ResolvedPart || seenWorksheets[e.ID] {
					return editRefusal("xlsx-opaque-cache-scope-unsupported", "unreferenced or wrong-target worksheet edge")
				}
				seenWorksheets[e.ID] = true
			default:
				// The exact sealed opaque profile contains worksheet edges only.
				// A supported relationship type is not evidence of inertness.
				return editRefusal("xlsx-opaque-cache-scope-unsupported", "extra workbook dependency")
			}
		} else if e.Source != "" {
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "outgoing semantic dependency")
		}
	}
	if workbookRootEdges != 1 {
		return editRefusal("xlsx-opaque-cache-scope-unsupported", "exactly one root workbook edge required")
	}
	if len(seenWorksheets) != len(worksheetIDs) {
		return editRefusal("xlsx-opaque-cache-scope-unsupported", "missing worksheet edge")
	}
	if len(seen) != 2 {
		return editRefusal("xlsx-opaque-cache-scope-unsupported", "missing inert root cache edge")
	}
	parts := map[string]string{}
	for _, p := range g.Parts {
		parts[p.Name] = p.ContentType
	}
	for _, want := range expected {
		if parts[want.part] != want.contentType {
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "cache content type differs")
		}
	}
	// Graph reports the effective content type. Require explicit reviewed
	// overrides, not an extension default that happens to resolve identically.
	contentTypes, _, err := s.pkg.Part(packaging.ContentTypesPath)
	if err != nil {
		return err
	}
	decoder := xml.NewDecoder(bytes.NewReader(contentTypes))
	overrides := map[string]string{}
	for {
		token, e := decoder.Token()
		if e != nil {
			if errors.Is(e, io.EOF) {
				break
			}
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "invalid content-type registry XML")
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "Override" || start.Name.Space != "http://schemas.openxmlformats.org/package/2006/content-types" {
			continue
		}
		var name, typ string
		for _, a := range start.Attr {
			if a.Name.Space != "" {
				return editRefusal("xlsx-opaque-cache-scope-unsupported", "unknown override attribute")
			}
			switch a.Name.Local {
			case "PartName":
				name = a.Value
			case "ContentType":
				typ = a.Value
			}
		}
		if _, dup := overrides[name]; dup {
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "duplicate cache override")
		}
		overrides[name] = typ
	}
	for _, want := range expected {
		if overrides["/"+want.part] != want.contentType {
			return editRefusal("xlsx-opaque-cache-scope-unsupported", "cache override missing or altered")
		}
	}
	// Deterministically examine both named parts, never infer from their bytes.
	names := make([]string, 0, len(expected))
	for _, want := range expected {
		names = append(names, want.part)
	}
	sort.Strings(names)
	for _, part := range names {
		if _, _, err := s.pkg.Part(part); err != nil {
			return err
		}
	}
	return nil
}
