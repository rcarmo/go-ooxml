package packaging_test

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const (
	profileCT   = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
	profileRels = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
	profileDoc  = `<?xml version="1.0" encoding="UTF-8"?><document>Alpha</document>`
)

func envelopeProfileSample(ct, rels, name string) []byte {
	data, err := packaging.WriteZIP32([]packaging.ZIP32Entry{{Name: "[Content_Types].xml", Data: []byte(ct)}, {Name: "_rels/.rels", Data: []byte(rels)}, {Name: name, Data: []byte(profileDoc)}})
	if err != nil {
		panic(err)
	}
	return data
}

func TestEnvelopeProfileNamedRefusalsAndCustody(t *testing.T) {
	for _, tc := range []struct{ name, ct, rels, part, reason string }{
		{"escaped part", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/%66oo.xml" ContentType="application/xml"/></Types>`, replaceTarget(profileRels, "word/%66oo.xml"), "word/%66oo.xml", "opc-part-name-invalid"},
		{"escaped target", profileCT, replaceTarget(profileRels, "word/%66oo.xml"), "word/document.xml", "opc-target-invalid"},
		{"duplicate default", replaceXMLDefault(profileCT), profileRels, "word/document.xml", "opc-content-types-invalid"},
		{"missing target", profileCT, replaceTarget(profileRels, "word/missing.xml"), "word/document.xml", "opc-relationship-target-missing"},
		{"duplicate rId", profileCT, duplicateRelation(profileRels), "word/document.xml", "opc-relationship-duplicate"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data := envelopeProfileSample(tc.ct, tc.rels, tc.part)
			original := bytes.Clone(data)
			model, err := packaging.OpenEnvelope(data)
			var refusal *packaging.ProfileError
			if model != nil || !errors.As(err, &refusal) || refusal.Reason != tc.reason {
				t.Fatalf("model=%v err=%v, want %s", model, err, tc.reason)
			}
			if !bytes.Equal(data, original) {
				t.Fatal("caller bytes changed")
			}
		})
	}
	data := envelopeProfileSample(profileCT, profileRels, "word/document.xml")
	original := bytes.Clone(data)
	model, err := packaging.OpenEnvelope(data)
	if err != nil {
		t.Fatal(err)
	}
	data[0] = 0
	root := t.TempDir()
	dest := filepath.Join(root, "out.docx")
	if err := os.WriteFile(dest, original, 0600); err != nil {
		t.Fatal(err)
	}
	if err := model.DeletePart("word/document.xml"); err != nil {
		t.Fatal(err)
	}
	err = model.SaveAs(dest)
	var refusal *packaging.ProfileError
	if !errors.As(err, &refusal) || refusal.Reason != "opc-relationship-target-missing" {
		t.Fatalf("invalid target save: %v", err)
	}
	// Invalid graph validation precedes even destination-directory I/O;
	// a verification failure after writing a temporary file is not enough.
	err = model.SaveAs(filepath.Join(root, "absent-directory", "out.docx"))
	if !errors.As(err, &refusal) || refusal.Reason != "opc-relationship-target-missing" {
		t.Fatalf("save did not preflight before destination I/O: %v", err)
	}
	read, err := os.ReadFile(dest)
	if err != nil || !bytes.Equal(read, original) {
		t.Fatal("destination changed", err)
	}
	fresh, err := packaging.OpenEnvelope(original)
	if err != nil {
		t.Fatal(err)
	}
	link := dest + ".link"
	if err := os.Symlink(dest, link); err != nil {
		t.Fatal(err)
	}
	err = fresh.SaveAs(link)
	if !errors.As(err, &refusal) || refusal.Reason != "opc-symlink-destination" {
		t.Fatalf("symlink refusal: %v", err)
	}
	read, err = os.ReadFile(dest)
	if err != nil || !bytes.Equal(read, original) {
		t.Fatal("symlink target changed", err)
	}
	gotLink, err := os.Readlink(link)
	if err != nil || gotLink != dest {
		t.Fatal("link changed", gotLink, err)
	}
	entries, err := os.ReadDir(root)
	if err != nil || len(entries) != 2 {
		t.Fatal("temporary delivery files remain", entries, err)
	}
	if _, err := packaging.OpenEnvelope(original); err != nil {
		t.Fatal("original archive corrupted", err)
	}
}

func replaceTarget(source, target string) string {
	return strings.Replace(source, `Target="word/document.xml"`, `Target="`+target+`"`, 1)
}
func replaceXMLDefault(source string) string {
	return strings.Replace(source, `Default Extension="xml"`, `Default Extension="rels"`, 1)
}
func duplicateRelation(source string) string {
	start := strings.Index(source, "<Relationship ")
	end := strings.Index(source[start:], "/>") + start + 2
	return source[:end] + source[start:end] + source[end:]
}
