package losslessxml

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strings"
	"unicode/utf8"
)

// encoding/xml rejects supplementary NameStartChar even though XML 1.0 fifth
// edition permits it. Replace only runes in markup names with unique four-byte
// ASCII aliases in the decoder's private copy; keep source byte positions and
// restore QNames before namespace expansion. Source text/attribute values and
// the document's immutable original bytes are never transformed.
func decoderNameCompatibility(source []byte) ([]byte, func(xml.Name) xml.Name, error) {
	if !hasSupplementaryNameRune(source) {
		return source, func(n xml.Name) xml.Name { return n }, nil
	}
	copyOfSource := bytes.Clone(source)
	aliases := map[rune]string{}
	reverse := map[string]rune{}
	next := 0
	aliasFor := func(r rune) (string, error) {
		if alias, ok := aliases[r]; ok {
			return alias, nil
		}
		for next < 26*26*26 {
			n := next
			next++
			candidate := string([]byte{'Q', byte('A' + n/(26*26)), byte('A' + n/26%26), byte('A' + n%26)})
			if bytes.Contains(source, []byte(candidate)) {
				continue
			}
			aliases[r] = candidate
			reverse[candidate] = r
			return candidate, nil
		}
		return "", fmt.Errorf("too many supplementary XML name components")
	}
	for at := 0; at < len(copyOfSource); {
		start := bytes.IndexByte(copyOfSource[at:], '<')
		if start < 0 {
			break
		}
		start += at
		if start+1 >= len(copyOfSource) {
			break
		}
		if bytes.HasPrefix(copyOfSource[start:], []byte("<!--")) {
			end := bytes.Index(copyOfSource[start+4:], []byte("-->"))
			if end < 0 {
				break
			}
			at = start + 4 + end + 3
			continue
		}
		if bytes.HasPrefix(copyOfSource[start:], []byte("<![CDATA[")) {
			end := bytes.Index(copyOfSource[start+9:], []byte("]]>"))
			if end < 0 {
				break
			}
			at = start + 9 + end + 3
			continue
		}
		if copyOfSource[start+1] == '!' || copyOfSource[start+1] == '?' {
			at = start + 2
			continue
		}
		quoted := byte(0)
		i := start + 1
		for i < len(copyOfSource) {
			b := copyOfSource[i]
			if quoted != 0 {
				if b == quoted {
					quoted = 0
				}
				i++
				continue
			}
			if b == '\'' || b == '"' {
				quoted = b
				i++
				continue
			}
			if b == '>' {
				i++
				break
			}
			if b < 0x80 {
				i++
				continue
			}
			r, width := utf8.DecodeRune(copyOfSource[i:])
			if r > 0xffff {
				alias, err := aliasFor(r)
				if err != nil {
					return nil, nil, err
				}
				copy(copyOfSource[i:i+width], alias)
			}
			i += width
		}
		at = i
	}
	restore := func(name xml.Name) xml.Name {
		replace := func(s string) string {
			for alias, r := range reverse {
				s = strings.ReplaceAll(s, alias, string(r))
			}
			return s
		}
		name.Space = replace(name.Space)
		name.Local = replace(name.Local)
		return name
	}
	return copyOfSource, restore, nil
}
func hasSupplementaryNameRune(source []byte) bool {
	for _, r := range string(source) {
		if r > 0xffff {
			return true
		}
	}
	return false
}
