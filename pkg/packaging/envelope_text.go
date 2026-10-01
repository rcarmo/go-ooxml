package packaging

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
)

// ReplaceXMLText changes one uniquely matching plain XML text leaf in an
// existing member. It retains UTF-8 and BOM-marked UTF-16 encodings and the
// original declaration spelling. Other XML and package members stay opaque.
// The new value is escaped by the production lexical editor, not the caller.
func (e *Envelope) ReplaceXMLText(part, old, replacement string) error {
	if e == nil || old == "" || !utf8.ValidString(replacement) {
		return fmt.Errorf("invalid XML text replacement")
	}
	original, err := e.Part(part)
	if err != nil {
		return err
	}
	decoded, err := decodeXMLMember(original)
	if err != nil {
		return err
	}
	doc, err := losslessxml.Parse(decoded)
	if err != nil {
		return err
	}
	var matches []losslessxml.TextEdit
	for _, element := range doc.Elements() {
		if value, plain := element.Text(); plain && value == old {
			matches = append(matches, losslessxml.TextEdit{Target: element, Text: replacement})
		}
	}
	if len(matches) != 1 {
		return &ProfileError{Reason: "opc-xml-member-invalid", Part: part, Cause: fmt.Errorf("XML text leaf not unique")}
	}
	changed, err := doc.ReplaceText(matches)
	if err != nil {
		return err
	}
	bom := byte(0)
	if bytes.HasPrefix(original, []byte{0xff, 0xfe}) {
		bom = 1
	} else if bytes.HasPrefix(original, []byte{0xfe, 0xff}) {
		bom = 2
	}
	if bom != 0 {
		// Restore the original declaration before re-encoding; decodeXMLMember
		// only normalises the encoding value for the UTF-8 lexical parser.
		declared := decodeUTF16Declaration(original, bom)
		if declared != "" {
			at := bytes.Index(changed, []byte("?>"))
			if at < 0 || !bytes.HasPrefix(changed, []byte("<?xml")) {
				return fmt.Errorf("translated XML declaration absent")
			}
			changed = append(append([]byte(nil), []byte(declared)...), changed[at+2:]...)
		}
		encoded := make([]byte, 0, len(changed)*2+2)
		if bom == 1 {
			encoded = append(encoded, 0xff, 0xfe)
		} else {
			encoded = append(encoded, 0xfe, 0xff)
		}
		for _, unit := range utf16.Encode([]rune(string(changed))) {
			var pair [2]byte
			if bom == 1 {
				binary.LittleEndian.PutUint16(pair[:], unit)
			} else {
				binary.BigEndian.PutUint16(pair[:], unit)
			}
			encoded = append(encoded, pair[:]...)
		}
		changed = encoded
	} else if bytes.HasPrefix(original, []byte{0xef, 0xbb, 0xbf}) {
		changed = append([]byte{0xef, 0xbb, 0xbf}, changed...)
	}
	// Verify the newly encoded member before staging it; no partial replacement.
	validated, err := decodeXMLMember(changed)
	if err != nil {
		return err
	}
	if _, err = losslessxml.Parse(validated); err != nil {
		return err
	}
	return e.SetPart(part, changed)
}

func decodeUTF16Declaration(raw []byte, bom byte) string {
	if len(raw) < 4 {
		return ""
	}
	units := make([]uint16, 0, (len(raw)-2)/2)
	for i := 2; i+1 < len(raw); i += 2 {
		if bom == 1 {
			units = append(units, binary.LittleEndian.Uint16(raw[i:i+2]))
		} else {
			units = append(units, binary.BigEndian.Uint16(raw[i:i+2]))
		}
	}
	decoded := string(utf16.Decode(units))
	if !strings.HasPrefix(decoded, "<?xml") {
		return ""
	}
	end := strings.Index(decoded, "?>")
	if end < 0 {
		return ""
	}
	// decodeXMLMember already checked the original BOM and declaration.
	// A UTF-8 xml.Decoder would reject encoding="UTF-16" here, so it cannot
	// serve as a second validation of the original UTF-16 declaration.
	return decoded[:end+2]
}
