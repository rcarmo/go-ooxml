package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"strings"
	"testing"
)

func admissionXMLArchive(t *testing.T) []byte {
	t.Helper()
	var b bytes.Buffer
	writer := zip.NewWriter(&b)
	h := &zip.FileHeader{Name: "a.xml", Method: zip.Deflate}
	out, err := writer.CreateHeader(h)
	if err != nil {
		t.Fatal(err)
	}
	if _, err = out.Write([]byte("<a>" + strings.Repeat(" ", 10000) + "</a>")); err != nil {
		t.Fatal(err)
	}
	if err = writer.Close(); err != nil {
		t.Fatal(err)
	}
	return b.Bytes()
}
func TestArchiveAdmissionBudgetsAndXMLFamily(t *testing.T) {
	source := admissionXMLArchive(t)
	original := bytes.Clone(source)
	zero := 0
	two := uint64(2)
	one := float64(1)
	tests := []struct {
		name   string
		limits ArchiveAdmissionLimits
		reason string
	}{
		{"zero members", ArchiveAdmissionLimits{MaxMembers: &zero}, "zip-too-many-entries"},
		{"member bytes", ArchiveAdmissionLimits{MaxMemberBytes: &two}, "zip-entry-too-large"},
		{"total bytes", ArchiveAdmissionLimits{MaxTotalBytes: &two}, "zip-total-too-large"},
		{"ratio", ArchiveAdmissionLimits{MaxRatio: &one}, "zip-compression-ratio-exceeded"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			entries, err := AdmitArchive(source, tc.limits)
			var refused *ProfileError
			if entries != nil || !errors.As(err, &refused) || refused.Reason != tc.reason {
				t.Fatalf("budget refusal: %v", err)
			}
			if !bytes.Equal(source, original) {
				t.Fatal("caller mutated")
			}
		})
	}
	entries, err := AdmitArchive(source, ArchiveAdmissionLimits{})
	if err != nil || len(entries) != 1 || len(entries[0].Data) != 10007 {
		t.Fatalf("default admission: %v len=%d", err, len(entries))
	}
	// An unsafe XML suffix must refuse without leaking any already-read member.
	var b bytes.Buffer
	writer := zip.NewWriter(&b)
	for _, part := range []struct{ name, text string }{{"safe.bin", "opaque"}, {"a.xml", `<!DOCTYPE a><a/>`}} {
		out, e := writer.Create(part.name)
		if e != nil {
			t.Fatal(e)
		}
		if _, e = out.Write([]byte(part.text)); e != nil {
			t.Fatal(e)
		}
	}
	if err = writer.Close(); err != nil {
		t.Fatal(err)
	}
	unsafe := b.Bytes()
	retained := bytes.Clone(unsafe)
	output, e := AdmitArchive(unsafe, ArchiveAdmissionLimits{})
	var refused *ProfileError
	if output != nil || !errors.As(e, &refused) || refused.Reason != "opc-xml-member-invalid" || !bytes.Equal(unsafe, retained) {
		t.Fatalf("unsafe member produced partial output or changed caller: %v", e)
	}
}
