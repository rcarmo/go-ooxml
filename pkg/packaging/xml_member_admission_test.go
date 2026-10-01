package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"errors"
	"strings"
	"testing"
	"unicode/utf16"
)

func utf16XML(source string, little bool) []byte {
	var out []byte
	if little {
		out = []byte{0xff, 0xfe}
	} else {
		out = []byte{0xfe, 0xff}
	}
	for _, unit := range utf16.Encode([]rune(source)) {
		var pair [2]byte
		if little {
			binary.LittleEndian.PutUint16(pair[:], unit)
		} else {
			binary.BigEndian.PutUint16(pair[:], unit)
		}
		out = append(out, pair[:]...)
	}
	return out
}
func independentXMLMemberZIP(name string, content []byte) []byte {
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	w, _ := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Store})
	_, _ = w.Write(content)
	_ = zw.Close()
	return out.Bytes()
}
func TestReadXMLMembersAdmissionFamily(t *testing.T) {
	cases := []struct {
		name    string
		payload []byte
		valid   bool
	}{
		{"utf8-dtd", []byte(`<!DOCTYPE a [<!ENTITY e "text">]><a>&e;</a>`), false},
		{"utf16le-dtd", utf16XML(`<!DOCTYPE a><a/>`, true), false},
		{"malformed", []byte(`<broken>`), false},
		{"valid-utf16le", utf16XML(`<?xml version="1.0" encoding="UTF-16"?><a/>`, true), true},
		{"valid-utf16le-spaced", utf16XML("<?xml version='1.0'\tencoding = 'utf-16' standalone='yes'?><a/>", true), true},
		{"valid-utf16be", utf16XML(`<?xml version="1.0" encoding = "UTF-16"?><a>😀</a>`, false), true},
		{"conflicting-utf16le", utf16XML(`<?xml version="1.0" encoding="UTF-8"?><a/>`, true), false},
		{"conflicting-utf16be", utf16XML(`<?xml version='1.0' encoding = 'ISO-8859-1'?><a/>`, false), false},
		{"valid-utf8-bom", append([]byte{0xef, 0xbb, 0xbf}, []byte(`<a/>`)...), true},
		{"comment-cdata", []byte(`<a><!-- <!DOCTYPE a> --><![CDATA[&custom;]]></a>`), true},
		{"unpaired", []byte{0xff, 0xfe, 0x00, 0xd8}, false},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			archive := independentXMLMemberZIP("a.xml", tc.payload)
			original := bytes.Clone(archive)
			entries, err := ReadXMLMembers(archive, ZIP32Limits{})
			if tc.valid {
				if err != nil || len(entries) != 1 || !bytes.Equal(entries[0].Data, tc.payload) {
					t.Fatalf("valid read: %v entries=%v", err, len(entries))
				}
				entries[0].Data[0] ^= 0xff
				again, err := ReadXMLMembers(archive, ZIP32Limits{})
				if err != nil || !bytes.Equal(again[0].Data, tc.payload) {
					t.Fatal("returned payload aliases source")
				}
			} else {
				var refused *ProfileError
				if entries != nil || !errors.As(err, &refused) || refused.Reason != "opc-xml-member-invalid" {
					t.Fatalf("unsafe XML read: %v", err)
				}
			}
			if !bytes.Equal(archive, original) {
				t.Fatal("caller archive changed")
			}
		})
	}
	for _, name := range []string{"meta.rels", "a.XML"} {
		archive := independentXMLMemberZIP(name, []byte(`<r>`))
		if result, err := ReadXMLMembers(archive, ZIP32Limits{}); result != nil || err == nil {
			t.Errorf("unsafe %s accepted", name)
		}
	}
	if out, err := ReadXMLMembers(independentXMLMemberZIP("opaque.bin", []byte(strings.Repeat("x", 8))), ZIP32Limits{}); err != nil || len(out) != 1 {
		t.Fatalf("opaque payload: %v", err)
	}
}
