package acceptance

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// contract20OraclePlaceholderText decodes the sealed source XML independently
// of the production lossless-XML reader. It observes only direct shape-owned
// text bodies and the visible paragraph/run/field/break sequence.
func contract20OraclePlaceholderText(data []byte, paragraphSeparator string, kinds ...string) (string, error) {
	p := func(local string) xml.Name { return xml.Name{Space: packaging.NSPresentationML, Local: local} }
	a := func(local string) xml.Name { return xml.Name{Space: packaging.NSDrawingML, Local: local} }
	var stack []xml.Name
	shapeDepth, bodyDepth, paragraphDepth, textDepth := -1, -1, -1, -1
	matches := 0
	selected := false
	var paragraphs []string
	var line strings.Builder
	var text strings.Builder
	var found string
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			parentDepth := len(stack) - 1
			parent := xml.Name{}
			if parentDepth >= 0 {
				parent = stack[parentDepth]
			}
			stack = append(stack, v.Name)
			depth := len(stack) - 1
			switch {
			case v.Name == p("sp") && shapeDepth < 0:
				shapeDepth, bodyDepth, paragraphDepth, textDepth = depth, -1, -1, -1
				selected = false
				paragraphs = nil
			case shapeDepth >= 0 && v.Name == p("ph"):
				for _, attr := range v.Attr {
					if attr.Name.Space == "" && attr.Name.Local == "type" {
						for _, kind := range kinds {
							if attr.Value == kind {
								selected = true
							}
						}
					}
				}
			case shapeDepth >= 0 && v.Name == p("txBody") && parent == p("sp") && parentDepth == shapeDepth:
				bodyDepth = depth
			case bodyDepth >= 0 && v.Name == a("p") && parent == p("txBody") && parentDepth == bodyDepth:
				paragraphDepth = depth
				line.Reset()
			case paragraphDepth >= 0 && v.Name == a("br") && parent == a("p") && parentDepth == paragraphDepth:
				line.WriteByte('\n')
			case paragraphDepth >= 0 && v.Name == a("t") && (parent == a("r") || parent == a("fld")) && parentDepth == paragraphDepth+1:
				textDepth = depth
				text.Reset()
			}
		case xml.CharData:
			if textDepth == len(stack)-1 && textDepth >= 0 {
				text.Write([]byte(v))
			}
		case xml.EndElement:
			depth := len(stack) - 1
			if depth < 0 || stack[depth] != v.Name {
				return "", fmt.Errorf("oracle XML stack mismatch")
			}
			switch depth {
			case textDepth:
				line.WriteString(strings.ReplaceAll(strings.ReplaceAll(text.String(), "\r\n", "\n"), "\r", "\n"))
				textDepth = -1
			case paragraphDepth:
				paragraphs = append(paragraphs, line.String())
				paragraphDepth = -1
			case bodyDepth:
				bodyDepth = -1
			case shapeDepth:
				if selected {
					matches++
					found = strings.Join(paragraphs, paragraphSeparator)
				}
				shapeDepth = -1
			}
			stack = stack[:depth]
		}
	}
	if len(stack) != 0 || matches != 1 {
		return "", fmt.Errorf("oracle %v placeholder count %d", kinds, matches)
	}
	return found, nil
}
