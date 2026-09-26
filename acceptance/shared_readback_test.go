package acceptance

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"strconv"
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

// Contract readback facts are checked independently against the input payloads.
// The bounded four-fixture family has no arbitrary path/query interpretation.
func verifySharedFacts(f mutationFixture, readback map[string]any, members map[string][]byte) error {
	count := func(part, space, local string) (int, error) {
		d, err := losslessxml.Parse(members[part])
		if err != nil {
			return 0, err
		}
		n := 0
		for _, e := range d.Elements() {
			if e.Name() == (xml.Name{Space: space, Local: local}) {
				n++
			}
		}
		return n, nil
	}
	measured := map[string]any{}
	switch f.ID {
	case "title-and-subtitle.pptx":
		n, err := count("ppt/presentation.xml", packaging.NSPresentationML, "sldId")
		if err != nil {
			return err
		}
		measured = map[string]any{"slideCount": n, "slide1": map[string]any{"title": readback["title"], "subtitle": readback["subtitle"]}}
	case "present-placeholder.docx":
		text, ok := readback["currentBodyText"].(string)
		if !ok {
			return fmt.Errorf("body readback absent")
		}
		absent, ok := f.Facts["absentText"].(string)
		if !ok || absent == "" || strings.Contains(text, absent) {
			return fmt.Errorf("absent text fact differs")
		}
		measured = map[string]any{"bodyText": text, "absentText": absent}
	case "default-style.xlsx":
		d, err := losslessxml.Parse(members["xl/styles.xml"])
		if err != nil {
			return err
		}
		n := 0
		for _, e := range d.Elements() {
			if e.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "xf"}) {
				if p, ok := e.Parent(); ok && p.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "cellXfs"}) {
					n++
				}
			}
		}
		absent, ok := f.Facts["absentSheet"].(string)
		if !ok || absent == "" {
			return fmt.Errorf("absent sheet fact missing")
		}
		wb, err := losslessxml.Parse(members["xl/workbook.xml"])
		if err != nil {
			return err
		}
		for _, e := range wb.Elements() {
			if e.Name() == (xml.Name{Space: packaging.NSSpreadsheetML, Local: "sheet"}) {
				for _, a := range e.Attributes() {
					if a.Name == (xml.Name{Local: "name"}) && a.Value == absent {
						return fmt.Errorf("absent sheet exists")
					}
				}
			}
		}
		measured = map[string]any{"activeSheetCellA1": readback["A1"], "cellXfsCount": n, "absentSheet": absent}
	case "cross-sheet-cache.xlsx":
		input, err := strconv.ParseFloat(fmt.Sprint(readback["Input!A1"]), 64)
		if err != nil {
			return err
		}
		cached, err := strconv.ParseFloat(fmt.Sprint(readback["Calc!A1.cached"]), 64)
		if err != nil {
			return err
		}
		_, chain := members["xl/calcChain.xml"]
		measured = map[string]any{"Input!A1": input, "Calc!A1": map[string]any{"formula": "=" + fmt.Sprint(readback["Calc!A1.formula"]), "cachedValue": cached}, "calcChainPresent": chain}
	default:
		return fmt.Errorf("unknown fact fixture %s", f.ID)
	}
	a, _ := json.Marshal(measured)
	b, _ := json.Marshal(f.Facts)
	if !bytes.Equal(a, b) {
		return fmt.Errorf("fixture readback facts differ for %s: got %s want %s", f.ID, a, b)
	}
	return nil
}
