package packaging

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
)

func TestNewPackage(t *testing.T) {
	pkg := New()
	if pkg == nil {
		t.Fatal("New() returned nil")
	}
	if pkg.ContentTypes() == nil {
		t.Error("ContentTypes() is nil")
	}
	if !pkg.IsModified() {
		t.Error("new package should be marked as modified")
	}
}

func TestPackage_AddAndGetPart(t *testing.T) {
	pkg := New()

	tests := []struct {
		name        string
		uri         string
		contentType string
		content     []byte
	}{
		{
			name:        "XML part",
			uri:         "word/document.xml",
			contentType: ContentTypeWordDocument,
			content:     []byte(`<?xml version="1.0"?><document/>`),
		},
		{
			name:        "with leading slash",
			uri:         "/word/styles.xml",
			contentType: ContentTypeStyles,
			content:     []byte(`<?xml version="1.0"?><styles/>`),
		},
		{
			name:        "binary content",
			uri:         "word/media/image1.png",
			contentType: ContentTypePNG,
			content:     []byte{0x89, 0x50, 0x4E, 0x47}, // PNG magic bytes
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			part, err := pkg.AddPart(tt.uri, tt.contentType, tt.content)
			if err != nil {
				t.Fatalf("AddPart() error = %v", err)
			}
			if part == nil {
				t.Fatal("AddPart() returned nil part")
			}

			// Retrieve the part
			retrieved, err := pkg.GetPart(tt.uri)
			if err != nil {
				t.Fatalf("GetPart() error = %v", err)
			}

			content, err := retrieved.Content()
			if err != nil {
				t.Fatalf("Content() error = %v", err)
			}
			if !bytes.Equal(content, tt.content) {
				t.Errorf("Content mismatch: got %v, want %v", content, tt.content)
			}
			if retrieved.ContentType() != tt.contentType {
				t.Errorf("ContentType() = %q, want %q", retrieved.ContentType(), tt.contentType)
			}
		})
	}
}

func TestPackage_DeletePart(t *testing.T) {
	pkg := New()
	_, _ = pkg.AddPart("test.xml", ContentTypeXML, []byte("<test/>"))

	if !pkg.PartExists("test.xml") {
		t.Error("part should exist after add")
	}

	err := pkg.DeletePart("test.xml")
	if err != nil {
		t.Fatalf("DeletePart() error = %v", err)
	}

	if pkg.PartExists("test.xml") {
		t.Error("part should not exist after delete")
	}

	// Delete non-existent part
	err = pkg.DeletePart("nonexistent.xml")
	if err == nil {
		t.Error("DeletePart() should error for non-existent part")
	}
}

func TestPackage_Relationships(t *testing.T) {
	pkg := New()

	// Add package-level relationship
	rel1 := pkg.AddRelationship("", "word/document.xml", RelTypeOfficeDocument)
	if rel1.ID != "rId1" {
		t.Errorf("first relationship ID = %q, want rId1", rel1.ID)
	}

	// Add part-level relationship
	rel2 := pkg.AddRelationship("word/document.xml", "styles.xml", RelTypeStyles)
	if rel2.ID != "rId1" {
		t.Errorf("first part relationship ID = %q, want rId1", rel2.ID)
	}

	// Get relationships by type
	docRels := pkg.GetRelationshipsByType("", RelTypeOfficeDocument)
	if len(docRels) != 1 {
		t.Errorf("expected 1 document relationship, got %d", len(docRels))
	}

	styleRels := pkg.GetRelationshipsByType("word/document.xml", RelTypeStyles)
	if len(styleRels) != 1 {
		t.Errorf("expected 1 styles relationship, got %d", len(styleRels))
	}
}

func TestPackage_Parts(t *testing.T) {
	pkg := New()
	_, _ = pkg.AddPart("part1.xml", ContentTypeXML, []byte("<p1/>"))
	_, _ = pkg.AddPart("part2.xml", ContentTypeXML, []byte("<p2/>"))
	_, _ = pkg.AddPart("part3.xml", ContentTypeXML, []byte("<p3/>"))

	parts := pkg.Parts()
	if len(parts) != 3 {
		t.Errorf("Parts() returned %d parts, want 3", len(parts))
	}
}

func TestPackage_SaveAndOpen(t *testing.T) {
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test.docx")

	// Create and save package
	pkg := New()
	_, _ = pkg.AddPart("word/document.xml", ContentTypeWordDocument, []byte(`<?xml version="1.0"?><document><body/></document>`))
	pkg.AddRelationship("", "word/document.xml", RelTypeOfficeDocument)

	err := pkg.SaveAs(tmpFile)
	if err != nil {
		t.Fatalf("SaveAs() error = %v", err)
	}
	pkg.Close()

	// Verify file exists
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("saved file does not exist")
	}

	// Open and verify
	pkg2, err := Open(tmpFile)
	if err != nil {
		t.Fatalf("Open() error = %v", err)
	}
	defer pkg2.Close()

	// Check content types were preserved
	ct := pkg2.GetContentType("word/document.xml")
	if ct != ContentTypeWordDocument {
		t.Errorf("content type = %q, want %q", ct, ContentTypeWordDocument)
	}

	// Check relationships were preserved
	rels := pkg2.GetRelationshipsByType("", RelTypeOfficeDocument)
	if len(rels) != 1 {
		t.Errorf("expected 1 relationship, got %d", len(rels))
	}

	// Check part content
	part, err := pkg2.GetPart("word/document.xml")
	if err != nil {
		t.Fatalf("GetPart() error = %v", err)
	}
	content, _ := part.Content()
	if !bytes.Contains(content, []byte("<document>")) {
		t.Error("part content not preserved")
	}
}

func TestPackage_OpenBytes(t *testing.T) {
	// First create a valid package
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test.docx")

	pkg := New()
	_, _ = pkg.AddPart("word/document.xml", ContentTypeWordDocument, []byte(`<document/>`))
	pkg.AddRelationship("", "word/document.xml", RelTypeOfficeDocument)
	_ = pkg.SaveAs(tmpFile)
	pkg.Close()

	// Read the file
	data, err := os.ReadFile(tmpFile)
	if err != nil {
		t.Fatalf("ReadFile() error = %v", err)
	}

	// Open from bytes
	pkg2, err := OpenBytes(data)
	if err != nil {
		t.Fatalf("OpenBytes() error = %v", err)
	}
	defer pkg2.Close()

	if !pkg2.PartExists("word/document.xml") {
		t.Error("part not found after OpenBytes")
	}
}

func TestPackage_ClosedOperations(t *testing.T) {
	pkg := New()
	pkg.Close()

	// Operations on closed package should fail
	_, err := pkg.GetPart("test.xml")
	if err == nil {
		t.Error("GetPart on closed package should error")
	}

	_, err = pkg.AddPart("test.xml", ContentTypeXML, []byte{})
	if err == nil {
		t.Error("AddPart on closed package should error")
	}

	err = pkg.DeletePart("test.xml")
	if err == nil {
		t.Error("DeletePart on closed package should error")
	}
}

func TestPackage_WriteTo(t *testing.T) {
	pkg := New()
	_, _ = pkg.AddPart("word/document.xml", ContentTypeWordDocument, []byte(`<document/>`))
	pkg.AddRelationship("", "word/document.xml", RelTypeOfficeDocument)

	var buf bytes.Buffer
	err := pkg.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo() error = %v", err)
	}

	// Should be a valid ZIP
	if buf.Len() == 0 {
		t.Error("WriteTo produced empty output")
	}

	// ZIP files start with PK
	if !bytes.HasPrefix(buf.Bytes(), []byte("PK")) {
		t.Error("output is not a valid ZIP file")
	}
}

func TestPart_Stream(t *testing.T) {
	pkg := New()
	content := []byte("test content for streaming")
	part, _ := pkg.AddPart("test.txt", "text/plain", content)

	stream, err := part.Stream()
	if err != nil {
		t.Fatalf("Stream() error = %v", err)
	}
	defer stream.Close()

	buf := make([]byte, len(content))
	n, err := stream.Read(buf)
	if err != nil {
		t.Fatalf("Read() error = %v", err)
	}
	if n != len(content) {
		t.Errorf("Read() = %d bytes, want %d", n, len(content))
	}
	if !bytes.Equal(buf, content) {
		t.Errorf("content mismatch")
	}
}

func TestPart_SetContent(t *testing.T) {
	pkg := New()
	part, _ := pkg.AddPart("test.xml", ContentTypeXML, []byte("original"))

	newContent := []byte("updated content")
	err := part.SetContent(newContent)
	if err != nil {
		t.Fatalf("SetContent() error = %v", err)
	}

	got, _ := part.Content()
	if !bytes.Equal(got, newContent) {
		t.Errorf("Content() = %q, want %q", got, newContent)
	}

	if !part.IsModified() {
		t.Error("part should be marked as modified")
	}
}

func TestPart_Size(t *testing.T) {
	pkg := New()
	content := []byte("12345678901234567890") // 20 bytes
	part, _ := pkg.AddPart("test.txt", "text/plain", content)

	if part.Size() != 20 {
		t.Errorf("Size() = %d, want 20", part.Size())
	}
}

func TestPackage_CorePropertiesRoundTrip(t *testing.T) {
	pkg := New()

	props := &common.CoreProperties{
		Title:          "Test Title",
		Creator:        "Test Author",
		Subject:        "Test Subject",
		Description:    "Test Description",
		Keywords:       "one;two",
		Category:       "Category",
		Language:       "en-US",
		ContentStatus:  "Draft",
		Identifier:     "urn:example:123",
		LastModifiedBy: "Reviewer",
		Revision:       "2",
		Version:        "1.0",
		Created:        &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-01-30T00:00:00Z"},
		Modified:       &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-01-30T01:00:00Z"},
		LastPrinted:    &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-01-30T02:00:00Z"},
	}

	if err := pkg.SetCoreProperties(props); err != nil {
		t.Fatalf("SetCoreProperties() error = %v", err)
	}

	got, err := pkg.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties() error = %v", err)
	}

	if got.Title != props.Title {
		t.Errorf("Title = %q, want %q", got.Title, props.Title)
	}
	if got.Creator != props.Creator {
		t.Errorf("Creator = %q, want %q", got.Creator, props.Creator)
	}
	if got.Subject != props.Subject {
		t.Errorf("Subject = %q, want %q", got.Subject, props.Subject)
	}
	if got.Description != props.Description {
		t.Errorf("Description = %q, want %q", got.Description, props.Description)
	}
	if got.Keywords != props.Keywords {
		t.Errorf("Keywords = %q, want %q", got.Keywords, props.Keywords)
	}
	if got.Category != props.Category {
		t.Errorf("Category = %q, want %q", got.Category, props.Category)
	}
	if got.Language != props.Language {
		t.Errorf("Language = %q, want %q", got.Language, props.Language)
	}
	if got.ContentStatus != props.ContentStatus {
		t.Errorf("ContentStatus = %q, want %q", got.ContentStatus, props.ContentStatus)
	}
	if got.Identifier != props.Identifier {
		t.Errorf("Identifier = %q, want %q", got.Identifier, props.Identifier)
	}
	if got.LastModifiedBy != props.LastModifiedBy {
		t.Errorf("LastModifiedBy = %q, want %q", got.LastModifiedBy, props.LastModifiedBy)
	}
	if got.Revision != props.Revision {
		t.Errorf("Revision = %q, want %q", got.Revision, props.Revision)
	}
	if got.Version != props.Version {
		t.Errorf("Version = %q, want %q", got.Version, props.Version)
	}
	if got.Created == nil || got.Created.Value != props.Created.Value {
		t.Errorf("Created = %v, want %v", got.Created, props.Created)
	}
	if got.Modified == nil || got.Modified.Value != props.Modified.Value {
		t.Errorf("Modified = %v, want %v", got.Modified, props.Modified)
	}
	if got.LastPrinted == nil || got.LastPrinted.Value != props.LastPrinted.Value {
		t.Errorf("LastPrinted = %v, want %v", got.LastPrinted, props.LastPrinted)
	}

	rels := pkg.GetRelationshipsByType("", RelTypeCoreProps)
	if len(rels) != 1 {
		t.Fatalf("expected 1 core properties relationship, got %d", len(rels))
	}
	if rels[0].Target != CorePropertiesPath {
		t.Errorf("core properties target = %q, want %q", rels[0].Target, CorePropertiesPath)
	}
}

func TestPackage_CoreProperties_Default(t *testing.T) {
	pkg := New()

	props, err := pkg.CoreProperties()
	if err != nil {
		t.Fatalf("CoreProperties() error = %v", err)
	}
	if props == nil {
		t.Fatal("CoreProperties() returned nil")
	}
	if props.Title != "" || props.Creator != "" {
		t.Errorf("expected empty default core properties, got title=%q creator=%q", props.Title, props.Creator)
	}
}
