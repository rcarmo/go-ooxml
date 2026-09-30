package xmlsnapshot

import (
	"fmt"
	"strings"
	"unicode/utf8"
)

// EscapeText escapes an XML 1.0 character value for element text. Literal CR
// must be a reference to survive XML line-ending normalization on reparse.
func EscapeText(value string) (string, error) { return escape(value, false) }

// EscapeAttribute escapes an XML 1.0 character value for a double-quoted
// attribute, preserving TAB/LF/CR across attribute-value normalization.
func EscapeAttribute(value string) (string, error) { return escape(value, true) }

// EscapeError identifies an invalid XML character or encoding. No partial
// escaped value is returned when a refusal occurs.
type EscapeError struct {
	Category string
	Cause    error
}

func (e *EscapeError) Error() string { return e.Cause.Error() }
func (e *EscapeError) Unwrap() error { return e.Cause }

func escape(value string, attribute bool) (string, error) {
	if !utf8.ValidString(value) {
		return "", &EscapeError{CategoryInvalidCharacter, fmt.Errorf("invalid UTF-8 XML value")}
	}
	var out strings.Builder
	for _, r := range value {
		if !(r == 9 || r == 10 || r == 13 || r >= 0x20 && r <= 0xd7ff || r >= 0xe000 && r <= 0xfffd || r >= 0x10000 && r <= 0x10ffff) {
			return "", &EscapeError{CategoryInvalidCharacter, fmt.Errorf("invalid XML character U+%04X", r)}
		}
		switch r {
		case '&':
			out.WriteString("&amp;")
		case '<':
			out.WriteString("&lt;")
		case '>':
			out.WriteString("&gt;")
		case '\'':
			if attribute {
				out.WriteString("&apos;")
			} else {
				out.WriteRune(r)
			}
		case '"':
			if attribute {
				out.WriteString("&quot;")
			} else {
				out.WriteRune(r)
			}
		case '\r':
			out.WriteString("&#13;")
		case '\n':
			if attribute {
				out.WriteString("&#10;")
			} else {
				out.WriteRune(r)
			}
		case '\t':
			if attribute {
				out.WriteString("&#9;")
			} else {
				out.WriteRune(r)
			}
		default:
			out.WriteRune(r)
		}
	}
	return out.String(), nil
}
