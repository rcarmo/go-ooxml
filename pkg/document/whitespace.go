package document

import (
	"encoding/xml"
	"strings"
	"unicode"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

func whitespaceAttributes(edits []losslessxml.TextEdit) []losslessxml.AttributeEdit {
	out := []losslessxml.AttributeEdit{}
	seen := map[losslessxml.Element]bool{}
	for _, edit := range edits {
		if seen[edit.Target] {
			continue
		}
		seen[edit.Target] = true
		old, _ := edit.Target.Text()
		if old == edit.Text {
			continue
		}
		significant := strings.TrimFunc(edit.Text, unicode.IsSpace) != edit.Text || strings.Contains(edit.Text, "  ")
		if !significant {
			continue
		}
		out = append(out, losslessxml.AttributeEdit{Target: edit.Target, Name: xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}, Value: "preserve"})
	}
	return out
}
