package packaging

import (
	"bytes"
	"encoding/binary"
	"strings"
	"testing"
	"unicode/utf16"
)

func textEnvelope(utf16LE bool) ([]byte, error) {
	xmlPart := []byte(`<?xml version="1.0" encoding="UTF-8"?><document>Alpha</document>`)
	if utf16LE {
		xmlPart = []byte{0xff, 0xfe}
		for _, u := range utf16.Encode([]rune(`<?xml version='1.0' encoding = 'UTF-16'?><document>Alpha</document>`)) {
			xmlPart = binary.LittleEndian.AppendUint16(xmlPart, u)
		}
	}
	return WriteZIP32([]ZIP32Entry{{Name: ContentTypesPath, Data: []byte(`<Types xmlns="` + NSContentTypes + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)}, {Name: PackageRelsPath, Data: []byte(`<Relationships xmlns="` + NSRelationships + `"><Relationship Id="rId1" Type="` + RelTypeOfficeDocument + `" Target="word/document.xml"/></Relationships>`)}, {Name: "word/document.xml", Data: xmlPart}})
}

func TestEnvelopeReplaceXMLTextNativeBatch(t *testing.T) {
	for _, tc := range []struct {
		name string
		wide bool
	}{{"utf8", false}, {"utf16le", true}} {
		t.Run(tc.name, func(t *testing.T) {
			input, err := textEnvelope(tc.wide)
			if err != nil {
				t.Fatal(err)
			}
			retained := bytes.Clone(input)
			e, err := OpenEnvelope(input)
			if err != nil {
				t.Fatal(err)
			}
			if err = e.ReplaceXMLText("word/document.xml", "Alpha", "Beta <value>"); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(input, retained) {
				t.Fatal("caller mutated")
			}
			saved, err := e.Bytes()
			if err != nil {
				t.Fatal(err)
			}
			again, err := OpenEnvelope(saved)
			if err != nil {
				t.Fatal(err)
			}
			part, err := again.Part("word/document.xml")
			if err != nil {
				t.Fatal(err)
			}
			if tc.wide {
				if !bytes.HasPrefix(part, []byte{0xff, 0xfe}) {
					t.Fatal("UTF-16 BOM lost")
				}
				units := make([]uint16, (len(part)-2)/2)
				for i := range units {
					units[i] = binary.LittleEndian.Uint16(part[i*2+2 : i*2+4])
				}
				text := string(utf16.Decode(units))
				if !strings.Contains(text, "encoding = 'UTF-16'") || !strings.Contains(text, "Beta &lt;value&gt;") {
					t.Fatalf("UTF-16 declaration/text: %q", text)
				}
			} else if !bytes.Contains(part, []byte("Beta &lt;value&gt;")) {
				t.Fatalf("text: %s", part)
			}
		})
	}
	input, err := textEnvelope(false)
	if err != nil {
		t.Fatal(err)
	}
	e, err := OpenEnvelope(input)
	if err != nil {
		t.Fatal(err)
	}
	if err = e.ReplaceXMLText("word/document.xml", "missing", "Beta"); err == nil {
		t.Fatal("missing leaf accepted")
	}
	saved, err := e.Bytes()
	if err != nil || !bytes.Equal(saved, input) {
		t.Fatal("refusal changed archive", err)
	}
}
