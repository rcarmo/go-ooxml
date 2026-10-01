package packaging

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"regexp"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/pkg/xmlsnapshot"
)

// ReadXMLMembers is an additive bounded raw-ZIP profile. It validates every
// .xml and .rels payload before returning any member, without requiring OPC
// registries. The input and returned entries never share payload storage.
func ReadXMLMembers(source []byte, limits ZIP32Limits) ([]ZIP32Entry, error) {
	entries, err := ReadZIP32(source, limits)
	if err != nil {
		return nil, err
	}
	for _, entry := range entries {
		lower := strings.ToLower(entry.Name)
		if !strings.HasSuffix(lower, ".xml") && !strings.HasSuffix(lower, ".rels") {
			continue
		}
		decoded, err := decodeXMLMember(entry.Data)
		if err == nil {
			_, err = xmlsnapshot.ParseLexical(decoded, xmlsnapshot.LexicalLimits{})
		}
		if err != nil {
			return nil, &ProfileError{Reason: "opc-xml-member-invalid", Part: entry.Name, Cause: err}
		}
	}
	return entries, nil
}

// decodeXMLMember accepts UTF-8 (optionally BOM) and BOM-marked UTF-16 in
// either byte order. UTF-16 decoding rejects unpaired surrogates rather than
// silently substituting U+FFFD. The converted buffer is private and temporary;
// callers still receive exact original member bytes from ReadZIP32.
func decodeXMLMember(source []byte) ([]byte, error) {
	if bytes.HasPrefix(source, []byte{0xef, 0xbb, 0xbf}) {
		return bytes.Clone(source[3:]), nil
	}
	little, offset := false, 0
	switch {
	case bytes.HasPrefix(source, []byte{0xff, 0xfe}):
		little, offset = true, 2
	case bytes.HasPrefix(source, []byte{0xfe, 0xff}):
		offset = 2
	default:
		return bytes.Clone(source), nil
	}
	if (len(source)-offset)%2 != 0 {
		return nil, fmt.Errorf("odd UTF-16 XML length")
	}
	units := make([]uint16, (len(source)-offset)/2)
	for i := range units {
		pair := source[offset+2*i : offset+2*i+2]
		if little {
			units[i] = binary.LittleEndian.Uint16(pair)
		} else {
			units[i] = binary.BigEndian.Uint16(pair)
		}
	}
	var text strings.Builder
	for i := 0; i < len(units); i++ {
		r := rune(units[i])
		if utf16.IsSurrogate(r) {
			if i+1 >= len(units) {
				return nil, fmt.Errorf("unpaired UTF-16 XML surrogate")
			}
			r = utf16.DecodeRune(r, rune(units[i+1]))
			if r == utf8.RuneError {
				return nil, fmt.Errorf("unpaired UTF-16 XML surrogate")
			}
			i++
		}
		text.WriteRune(r)
	}
	decoded := text.String()
	// Inspect only a leading XML declaration. The parser subsequently validates
	// its grammar; this conversion changes the encoding value, not whitespace,
	// quotes, attribute ordering or any other declaration tokens.
	if strings.HasPrefix(decoded, "<?xml") && len(decoded) > 5 && strings.ContainsAny(decoded[5:6], " \t\r\n") {
		end := strings.Index(decoded, "?>")
		if end < 0 {
			return nil, fmt.Errorf("unterminated UTF-16 XML declaration")
		}
		decl := decoded[:end]
		if match := xmlDeclarationEncoding.FindStringIndex(decl); match != nil {
			start := match[1]
			if start >= len(decl) || (decl[start] != '\'' && decl[start] != '"') {
				return nil, fmt.Errorf("unquoted UTF-16 XML encoding")
			}
			quote := decl[start]
			stop := strings.IndexByte(decl[start+1:], quote)
			if stop < 0 {
				return nil, fmt.Errorf("unterminated UTF-16 XML encoding")
			}
			stop += start + 1
			if !strings.EqualFold(decl[start+1:stop], "UTF-16") {
				return nil, fmt.Errorf("UTF-16 BOM conflicts with declared encoding")
			}
			decoded = decoded[:start+1] + "UTF-8" + decoded[stop:]
		}
	}
	return []byte(decoded), nil
}

var xmlDeclarationEncoding = regexp.MustCompile(`(?:^|[ \t\r\n])encoding[ \t\r\n]*=[ \t\r\n]*`)
