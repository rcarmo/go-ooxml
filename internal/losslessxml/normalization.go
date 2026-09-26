package losslessxml

import (
	"bytes"
	"encoding/xml"
	"fmt"
)

// encoding/xml normalises literal CR/CRLF but leaves literal attribute LF/TAB.
// XML 1.0 3.3.3 additionally maps these literal whitespace characters to spaces;
// numeric references must retain their referenced characters. Work on lexical
// values before entity decoding, never on the original source buffer.
func normalizeAttributeValues(startTag []byte, attributes []xml.Attr) ([]xml.Attr, error) {
	if len(attributes) == 0 {
		return attributes, nil
	}
	at := 1
	space := func(b byte) bool { return b == ' ' || b == '\t' || b == '\r' || b == '\n' }
	for at < len(startTag) && !space(startTag[at]) && startTag[at] != '>' && startTag[at] != '/' {
		at++
	}
	out := append([]xml.Attr(nil), attributes...)
	index := 0
	for at < len(startTag) {
		for at < len(startTag) && space(startTag[at]) {
			at++
		}
		if at >= len(startTag) || startTag[at] == '>' || startTag[at] == '/' {
			break
		}
		for at < len(startTag) && !space(startTag[at]) && startTag[at] != '=' {
			at++
		}
		for at < len(startTag) && space(startTag[at]) {
			at++
		}
		if at >= len(startTag) || startTag[at] != '=' {
			return nil, fmt.Errorf("attribute equals missing")
		}
		at++
		for at < len(startTag) && space(startTag[at]) {
			at++
		}
		if at >= len(startTag) || (startTag[at] != '\'' && startTag[at] != '"') {
			return nil, fmt.Errorf("attribute quote missing")
		}
		quote := startTag[at]
		at++
		begin := at
		for at < len(startTag) && startTag[at] != quote {
			at++
		}
		if at >= len(startTag) || index >= len(out) {
			return nil, fmt.Errorf("attribute lexical mapping differs")
		}
		raw := startTag[begin:at]
		at++
		if bytes.ContainsAny(raw, "\t\r\n") {
			value := make([]byte, 0, len(raw))
			for i := 0; i < len(raw); i++ {
				b := raw[i]
				if b == '\r' {
					if i+1 < len(raw) && raw[i+1] == '\n' {
						i++
					}
					b = ' '
				} else if b == '\n' || b == '\t' {
					b = ' '
				}
				value = append(value, b)
			}
			var fragment bytes.Buffer
			fragment.WriteString("<v a=")
			fragment.WriteByte(quote)
			fragment.Write(value)
			fragment.WriteByte(quote)
			fragment.WriteString("/>")
			token, err := xml.NewDecoder(bytes.NewReader(fragment.Bytes())).RawToken()
			if err != nil {
				return nil, err
			}
			e, ok := token.(xml.StartElement)
			if !ok || len(e.Attr) != 1 {
				return nil, fmt.Errorf("invalid normalised attribute")
			}
			out[index].Value = e.Attr[0].Value
		}
		index++
	}
	if index != len(out) {
		return nil, fmt.Errorf("attribute count differs from lexical map")
	}
	return out, nil
}
