package packaging

import "testing"

func TestContentTypesGetContentType(t *testing.T) {
	ct := NewContentTypes()
	ct.AddOverride("/word/document.xml", ContentTypeWordDocument)
	ct.AddOverride("/word/styles.xml", ContentTypeStyles)

	tests := []struct {
		uri    string
		expect string
	}{
		{"/word/document.xml", ContentTypeWordDocument},
		{"word/document.xml", ContentTypeWordDocument}, // without leading /
		{"/word/styles.xml", ContentTypeStyles},
		{"/word/unknown.xml", ContentTypeXML}, // falls back to extension default
		{"/image.png", ContentTypePNG},
		{"/image.jpeg", ContentTypeJPEG},
		{"/unknown.xyz", ""}, // unknown extension
	}

	for _, tt := range tests {
		got := ct.GetContentType(tt.uri)
		if got != tt.expect {
			t.Errorf("GetContentType(%q) = %q, want %q", tt.uri, got, tt.expect)
		}
	}
}

func TestContentTypesAddOverride(t *testing.T) {
	ct := NewContentTypes()

	ct.AddOverride("/word/document.xml", ContentTypeWordDocument)
	if len(ct.Overrides) != 1 {
		t.Errorf("len(Overrides) = %d, want 1", len(ct.Overrides))
	}

	// Adding same part should update, not duplicate
	ct.AddOverride("/word/document.xml", ContentTypeWordTemplate)
	if len(ct.Overrides) != 1 {
		t.Errorf("After update, len(Overrides) = %d, want 1", len(ct.Overrides))
	}
	if ct.GetContentType("/word/document.xml") != ContentTypeWordTemplate {
		t.Error("Override was not updated")
	}
}

func TestContentTypesAddDefault(t *testing.T) {
	ct := NewContentTypes()
	initialLen := len(ct.Defaults)

	// Add new extension
	ct.AddDefault("docx", "application/vnd.custom")
	if len(ct.Defaults) != initialLen+1 {
		t.Errorf("After add, len(Defaults) = %d, want %d", len(ct.Defaults), initialLen+1)
	}

	// Update existing extension (should not add duplicate)
	ct.AddDefault("png", "image/custom-png")
	if len(ct.Defaults) != initialLen+1 {
		t.Errorf("After update, len(Defaults) = %d, want %d", len(ct.Defaults), initialLen+1)
	}
}

func TestContentTypesRemoveOverride(t *testing.T) {
	ct := NewContentTypes()
	ct.AddOverride("/word/document.xml", ContentTypeWordDocument)

	if !ct.RemoveOverride("/word/document.xml") {
		t.Error("RemoveOverride returned false")
	}
	if len(ct.Overrides) != 0 {
		t.Errorf("After remove, len(Overrides) = %d, want 0", len(ct.Overrides))
	}
	if ct.RemoveOverride("/word/document.xml") {
		t.Error("Second RemoveOverride should return false")
	}
}

func TestContentTypesEnsureContentType(t *testing.T) {
	ct := NewContentTypes()

	// For XML file, default is XML - should not add override
	ct.EnsureContentType("/word/unknown.xml", ContentTypeXML)
	if len(ct.Overrides) != 0 {
		t.Error("EnsureContentType added unnecessary override for XML")
	}

	// For document, override needed
	ct.EnsureContentType("/word/document.xml", ContentTypeWordDocument)
	if len(ct.Overrides) != 1 {
		t.Error("EnsureContentType should have added override for Word document")
	}
}
