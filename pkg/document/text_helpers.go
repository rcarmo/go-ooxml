package document

import (
	"strconv"
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
		case *wml.Sym:
			sb.WriteRune(symToRune(v.Char))
		}
	}
	return sb.String()
}

func symToRune(char string) rune {
	if len(char) == 1 {
		return rune(char[0])
	}
	if len(char) == 0 {
		return 0
	}
	val, err := strconv.ParseUint(char, 16, 16)
	if err != nil {
		return 0
	}
	return rune(val)
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
		case *wml.Ins:
			sb.WriteString(textFromInlineContent(v.Content))
		case *wml.Del:
			for _, delElem := range v.Content {
				if run, ok := delElem.(*wml.R); ok {
					for _, runElem := range run.Content {
						if dt, ok := runElem.(*wml.DelText); ok {
							sb.WriteString(dt.Text)
						}
					}
				}
			}
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
		case *wml.Tbl:
			if sb.Len() > 0 {
				sb.WriteString("\n")
			}
			sb.WriteString(textFromTable(v))
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

func textFromTable(tbl *wml.Tbl) string {
	var sb strings.Builder
	for _, row := range tbl.Tr {
		for cellIndex, cell := range row.Tc {
			if cellIndex > 0 {
				sb.WriteString("\t")
			}
			sb.WriteString(textFromTableCell(cell))
		}
		if sb.Len() > 0 {
			sb.WriteString("\n")
		}
	}
	return strings.TrimSuffix(sb.String(), "\n")
}

func textFromTableCell(cell *wml.Tc) string {
	if cell == nil {
		return ""
	}
	var sb strings.Builder
	for i, elem := range cell.Content {
		if i > 0 {
			sb.WriteString("\n")
		}
		switch v := elem.(type) {
		case *wml.P:
			sb.WriteString(textFromParagraph(v))
		case *wml.Tbl:
			sb.WriteString(textFromTable(v))
		case *wml.Sdt:
			sb.WriteString(textFromSdt(v))
		}
	}
	return sb.String()
}
