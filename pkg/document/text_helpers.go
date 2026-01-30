package document

import (
	"strings"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

func textFromRun(r *wml.R) string {
	var sb strings.Builder
	for _, elem := range r.Content {
		switch v := elem.(type) {
		case *wml.T:
			sb.WriteString(v.Text)
		case *wml.Br:
			sb.WriteString("\n")
		case *wml.Tab:
			sb.WriteString("\t")
		}
	}
	return sb.String()
}

func textFromParagraph(p *wml.P) string {
	return textFromInlineContent(p.Content)
}

func textFromInlineContent(content []interface{}) string {
	var sb strings.Builder
	for _, elem := range content {
		switch v := elem.(type) {
		case *wml.R:
			sb.WriteString(textFromRun(v))
		case *wml.Hyperlink:
			sb.WriteString(textFromInlineContent(v.Content))
		case *wml.Sdt:
			sb.WriteString(textFromSdt(v))
		}
	}
	return sb.String()
}

func textFromSdt(sdt *wml.Sdt) string {
	if sdt.SdtContent == nil {
		return ""
	}
	var sb strings.Builder
	for _, elem := range sdt.SdtContent.Content {
		switch v := elem.(type) {
		case *wml.P:
			if sb.Len() > 0 {
				sb.WriteString("\n")
			}
			sb.WriteString(textFromParagraph(v))
		case *wml.R:
			sb.WriteString(textFromRun(v))
		case *wml.Hyperlink:
			sb.WriteString(textFromInlineContent(v.Content))
		case *wml.Sdt:
			sb.WriteString(textFromSdt(v))
		}
	}
	return sb.String()
}
