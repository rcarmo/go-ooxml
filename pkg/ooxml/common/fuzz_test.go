package common

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzCorePropertiesRoundTrip(f *testing.F) {
	f.Add("Title", "Author", "Subject")
	f.Add("Doc", "Creator", "Summary")

	f.Fuzz(func(t *testing.T, title, creator, subject string) {
		if len(title) > 128 || len(creator) > 128 || len(subject) > 256 {
			return
		}
		props := &CoreProperties{
			Title:   title,
			Creator: creator,
			Subject: subject,
		}
		data, err := utils.MarshalXMLWithHeader(props)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		var parsed CoreProperties
		if err := utils.UnmarshalXML(data, &parsed); err != nil {
			t.Fatalf("UnmarshalXML error: %v", err)
		}
		if parsed.Title != title || parsed.Creator != creator || parsed.Subject != subject {
			t.Fatalf("round-trip mismatch")
		}
	})
}

func FuzzSharedStringsRoundTrip(f *testing.F) {
	f.Add("Hello", "World")
	f.Add("Foo", "Bar")

	f.Fuzz(func(t *testing.T, a, b string) {
		if len(a) > 256 || len(b) > 256 {
			return
		}
		ss := &SharedStrings{}
		ss.AddString(a)
		ss.AddString(b)

		data, err := utils.MarshalXMLWithHeader(ss)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		var parsed SharedStrings
		if err := utils.UnmarshalXML(data, &parsed); err != nil {
			t.Fatalf("UnmarshalXML error: %v", err)
		}
		if parsed.GetString(0) != a && !strings.Contains(parsed.GetString(0), a) {
			t.Fatalf("shared string mismatch")
		}
	})
}
