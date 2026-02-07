package spreadsheet

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"sort"
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/sml"
)

type vmlCommentsRoot struct {
	XMLName xml.Name `xml:"xml"`
	Content string   `xml:",innerxml"`
}

func buildCommentsVML(comments *SheetComments) ([]byte, error) {
	if comments == nil || comments.comments == nil || comments.comments.CommentList == nil {
		return nil, nil
	}
	commentList := comments.comments.CommentList.Comment
	if len(commentList) == 0 {
		return nil, nil
	}
	ordered := append([]*sml.Comment(nil), commentList...)
	sort.SliceStable(ordered, func(i, j int) bool {
		return ordered[i].Ref < ordered[j].Ref
	})
	shapeIndex := 0
	for _, c := range ordered {
		if c.ShapeID == "" {
			c.ShapeID = "0"
		}
		shapeIndex++
	}
	var buf bytes.Buffer
	buf.WriteString(`<o:shapelayout xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" v:ext="edit">`)
	buf.WriteString(`<o:idmap v:ext="edit" data="1"/>`)
	buf.WriteString(`</o:shapelayout>`)
	buf.WriteString(`<v:shapetype xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">`)
	buf.WriteString(`<v:stroke joinstyle="miter"/>`)
	buf.WriteString(`<v:path gradientshapeok="t" o:connecttype="rect"/>`)
	buf.WriteString(`</v:shapetype>`)
	for i, c := range ordered {
		row, col := commentRowCol(c.Ref)
		if row < 0 || col < 0 {
			continue
		}
		shapeID := 1026 + i
		buf.WriteString(`<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" type="#_x0000_t202"`)
		buf.WriteString(` style="position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:144px;height:79px;z-index:1;visibility:hidden"`)
		buf.WriteString(` fillcolor="#ffffe1" o:insetmode="auto"`)
		buf.WriteString(fmt.Sprintf(` id="_x0000_s%d">`, shapeID))
		buf.WriteString(`<v:fill color2="#ffffe1"/>`)
		buf.WriteString(`<v:shadow color="black" obscured="t"/>`)
		buf.WriteString(`<v:path o:connecttype="none"/>`)
		buf.WriteString(`<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>`)
		buf.WriteString(`<x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="Note">`)
		buf.WriteString(`<x:MoveWithCells/>`)
		buf.WriteString(`<x:SizeWithCells/>`)
		buf.WriteString(`<x:AutoFill>False</x:AutoFill>`)
		buf.WriteString(fmt.Sprintf(`<x:Row>%d</x:Row>`, row))
		buf.WriteString(fmt.Sprintf(`<x:Column>%d</x:Column>`, col))
		buf.WriteString(`</x:ClientData>`)
		buf.WriteString(`</v:shape>`)
	}
	root := &vmlCommentsRoot{Content: buf.String()}
	data, err := xml.Marshal(root)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func commentRowCol(ref string) (int, int) {
	if ref == "" {
		return -1, -1
	}
	ref = strings.ToUpper(ref)
	var colStr strings.Builder
	var rowStr strings.Builder
	for _, r := range ref {
		if r >= 'A' && r <= 'Z' {
			colStr.WriteRune(r)
		} else if r >= '0' && r <= '9' {
			rowStr.WriteRune(r)
		}
	}
	if colStr.Len() == 0 || rowStr.Len() == 0 {
		return -1, -1
	}
	col := 0
	for _, r := range colStr.String() {
		col = col*26 + int(r-'A'+1)
	}
	row := 0
	for _, r := range rowStr.String() {
		row = row*10 + int(r-'0')
	}
	return row - 1, col - 1
}
