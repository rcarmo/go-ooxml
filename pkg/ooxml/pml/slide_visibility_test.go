package pml

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestSlideShowAttributeQualification(t *testing.T) {
	const ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
	cases := []struct {
		name, attrs string
		want        *bool
	}{
		{"default", "", nil},
		{"namespaced-only", `p:show="0"`, nil},
		{"foreign-only", `x:show="0"`, nil},
		{"unqualified-hidden", `show="0"`, boolPtr(false)},
		{"unqualified-visible", `show="1"`, boolPtr(true)},
		{"both-qualified-first", `p:show="1" show="0"`, boolPtr(false)},
		{"both-unqualified-first", `show="0" p:show="1"`, boolPtr(false)},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			input := `<p:sld xmlns:p="` + ns + `" xmlns:x="urn:test" ` + tc.attrs + `><p:cSld><p:spTree/></p:cSld></p:sld>`
			var slide Sld
			if err := xml.Unmarshal([]byte(input), &slide); err != nil {
				t.Fatal(err)
			}
			if (slide.Show == nil) != (tc.want == nil) || (slide.Show != nil && *slide.Show != *tc.want) {
				t.Fatalf("Show = %v, want %v", slide.Show, tc.want)
			}
			output, err := xml.Marshal(&slide)
			if err != nil {
				t.Fatal(err)
			}
			if strings.Contains(string(output), `p:show=`) || strings.Contains(string(output), `x:show=`) {
				t.Fatalf("qualified visibility serialized as CT_Slide show: %s", output)
			}
			var reopened Sld
			if err := xml.Unmarshal(output, &reopened); err != nil {
				t.Fatal(err)
			}
			if (reopened.Show == nil) != (tc.want == nil) || (reopened.Show != nil && *reopened.Show != *tc.want) {
				t.Fatalf("reopened Show = %v, want %v", reopened.Show, tc.want)
			}
		})
	}
}

func boolPtr(v bool) *bool { return &v }
