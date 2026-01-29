package utils

import (
	"bytes"
	"encoding/xml"
	"testing"
)

func TestXMLHeader(t *testing.T) {
	expected := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"
	if XMLHeader != expected {
		t.Errorf("XMLHeader = %q, want %q", XMLHeader, expected)
	}
}

func TestMarshalXMLWithHeader(t *testing.T) {
	type TestStruct struct {
		XMLName xml.Name `xml:"test"`
		Value   string   `xml:"value"`
	}

	tests := []struct {
		name    string
		input   interface{}
		wantErr bool
		check   func([]byte) bool
	}{
		{
			name:    "simple struct",
			input:   TestStruct{Value: "hello"},
			wantErr: false,
			check: func(b []byte) bool {
				return bytes.HasPrefix(b, []byte("<?xml")) && bytes.Contains(b, []byte("<test>"))
			},
		},
		{
			name:    "empty struct",
			input:   TestStruct{},
			wantErr: false,
			check: func(b []byte) bool {
				return bytes.HasPrefix(b, []byte("<?xml"))
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := MarshalXMLWithHeader(tt.input)
			if (err != nil) != tt.wantErr {
				t.Errorf("MarshalXMLWithHeader() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !tt.wantErr && !tt.check(got) {
				t.Errorf("MarshalXMLWithHeader() output check failed: %s", string(got))
			}
		})
	}
}

func TestMarshalXMLIndentWithHeader(t *testing.T) {
	type Inner struct {
		Text string `xml:"text"`
	}
	type TestStruct struct {
		XMLName xml.Name `xml:"test"`
		Inner   Inner    `xml:"inner"`
	}

	data, err := MarshalXMLIndentWithHeader(TestStruct{Inner: Inner{Text: "hello"}}, "", "  ")
	if err != nil {
		t.Fatalf("MarshalXMLIndentWithHeader() error = %v", err)
	}

	if !bytes.HasPrefix(data, []byte("<?xml")) {
		t.Error("missing XML header")
	}
	if !bytes.Contains(data, []byte("\n  ")) {
		t.Error("output not indented")
	}
}

func TestUnmarshalXML(t *testing.T) {
	type TestStruct struct {
		Value string `xml:"value"`
	}

	tests := []struct {
		name    string
		input   []byte
		want    string
		wantErr bool
	}{
		{
			name:    "normal XML",
			input:   []byte(`<root><value>test</value></root>`),
			want:    "test",
			wantErr: false,
		},
		{
			name:    "with BOM",
			input:   append([]byte{0xEF, 0xBB, 0xBF}, []byte(`<root><value>bom test</value></root>`)...),
			want:    "bom test",
			wantErr: false,
		},
		{
			name:    "invalid XML",
			input:   []byte(`<root><value>unclosed`),
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var result TestStruct
			err := UnmarshalXML(tt.input, &result)
			if (err != nil) != tt.wantErr {
				t.Errorf("UnmarshalXML() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !tt.wantErr && result.Value != tt.want {
				t.Errorf("UnmarshalXML() got = %q, want %q", result.Value, tt.want)
			}
		})
	}
}

func TestEscapeXMLText(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"hello", "hello"},
		{"<script>", "&lt;script&gt;"},
		{"a & b", "a &amp; b"},
		{`"quoted"`, "&#34;quoted&#34;"},
		{"line1\nline2", "line1&#xA;line2"}, // xml.EscapeText escapes newlines
		{"", ""},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			if got := EscapeXMLText(tt.input); got != tt.want {
				t.Errorf("EscapeXMLText(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}

func TestPointerHelpers(t *testing.T) {
	// BoolPtr
	bp := BoolPtr(true)
	if bp == nil || *bp != true {
		t.Error("BoolPtr(true) failed")
	}

	// IntPtr
	ip := IntPtr(42)
	if ip == nil || *ip != 42 {
		t.Error("IntPtr(42) failed")
	}

	// Int64Ptr
	i64p := Int64Ptr(12345678901234)
	if i64p == nil || *i64p != 12345678901234 {
		t.Error("Int64Ptr failed")
	}

	// StringPtr
	sp := StringPtr("test")
	if sp == nil || *sp != "test" {
		t.Error("StringPtr failed")
	}

	// Float64Ptr
	fp := Float64Ptr(3.14159)
	if fp == nil || *fp != 3.14159 {
		t.Error("Float64Ptr failed")
	}
}

func TestDerefHelpers(t *testing.T) {
	// DerefBool
	trueVal := true
	if DerefBool(&trueVal, false) != true {
		t.Error("DerefBool with value failed")
	}
	if DerefBool(nil, true) != true {
		t.Error("DerefBool with nil failed")
	}

	// DerefString
	strVal := "hello"
	if DerefString(&strVal, "default") != "hello" {
		t.Error("DerefString with value failed")
	}
	if DerefString(nil, "default") != "default" {
		t.Error("DerefString with nil failed")
	}

	// DerefInt
	intVal := 42
	if DerefInt(&intVal, 0) != 42 {
		t.Error("DerefInt with value failed")
	}
	if DerefInt(nil, 99) != 99 {
		t.Error("DerefInt with nil failed")
	}
}
