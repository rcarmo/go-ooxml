package acceptance

import (
	"encoding/xml"
	"fmt"
	"strings"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Readbacks are native checks of pinned inputs, not expected outputs or execution
// of the shared mutation scenarios. No Python or MCP participates here.
func sharedReadback(id string, data []byte, members map[string][]byte) (map[string]any, error) {
	switch id {
	case "present-placeholder.docx":
		s, err := document.OpenEditing(data, packaging.Limits{})
		if err != nil {
			return nil, err
		}
		o, err := s.Outline(document.CurrentView)
		if err != nil {
			return nil, err
		}
		var paragraphs []string
		for _, b := range o.Blocks {
			if b.Part == "word/document.xml" {
				paragraphs = append(paragraphs, b.Text)
			}
		}
		text := strings.Join(paragraphs, "\n")
		if text != "<Present>" {
			return nil, fmt.Errorf("body readback %q", text)
		}
		return map[string]any{"currentBodyText": text, "unreadable": o.Unreadable}, nil
	case "title-and-subtitle.pptx":
		s, err := presentation.OpenEditing(data, packaging.Limits{})
		if err != nil {
			return nil, err
		}
		if _, err = s.FindText("ppt/slides/slide1.xml", 2, "Original title"); err != nil {
			return nil, err
		}
		if _, err = s.FindText("ppt/slides/slide1.xml", 3, "Original subtitle"); err != nil {
			return nil, err
		}
		return map[string]any{"slideCount": 1, "title": "Original title", "subtitle": "Original subtitle"}, nil
	case "default-style.xlsx":
		value, formula, err := sharedCell(members["xl/worksheets/sheet1.xml"], "A1")
		if err != nil {
			return nil, err
		}
		if value != "before" || formula != "" {
			return nil, fmt.Errorf("default cell %q formula %q", value, formula)
		}
		return map[string]any{"A1": value, "formula": formula, "styles": "validated"}, nil
	case "cross-sheet-cache.xlsx":
		input, _, err := sharedCell(members["xl/worksheets/sheet1.xml"], "A1")
		if err != nil {
			return nil, err
		}
		cached, formula, err := sharedCell(members["xl/worksheets/sheet2.xml"], "A1")
		if err != nil {
			return nil, err
		}
		if input != "1" || cached != "2" || formula != "Input!A1*2" {
			return nil, fmt.Errorf("cross-sheet readback %q %q %q", input, cached, formula)
		}
		if _, ok := members["xl/calcChain.xml"]; ok {
			return nil, fmt.Errorf("unexpected calcChain")
		}
		return map[string]any{"Input!A1": input, "Calc!A1.cached": cached, "Calc!A1.formula": formula, "calculationPerformed": false}, nil
	default:
		return nil, fmt.Errorf("unknown shared fixture %s", id)
	}
}
func sharedCell(data []byte, address string) (string, string, error) {
	d, err := losslessxml.Parse(data)
	if err != nil {
		return "", "", err
	}
	var cells []losslessxml.Element
	for _, e := range d.Elements() {
		if e.Name() != (xml.Name{Space: packaging.NSSpreadsheetML, Local: "c"}) {
			continue
		}
		for _, a := range e.Attributes() {
			if a.Name == (xml.Name{Local: "r"}) && a.Value == address {
				cells = append(cells, e)
			}
		}
	}
	if len(cells) != 1 {
		return "", "", fmt.Errorf("ambiguous/absent cell %s", address)
	}
	value, formula := "", ""
	for _, e := range d.Elements() {
		belongs := false
		for p, ok := e.Parent(); ok; p, ok = p.Parent() {
			if p == cells[0] {
				belongs = true
				break
			}
		}
		if !belongs || e.Name().Space != packaging.NSSpreadsheetML {
			continue
		}
		text, leaf := e.Text()
		if !leaf {
			continue
		}
		switch e.Name().Local {
		case "v", "t":
			value += text
		case "f":
			formula += text
		}
	}
	return value, formula, nil
}
