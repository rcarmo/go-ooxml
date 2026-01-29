package packaging

import "testing"

func TestRelationshipsPathForPart(t *testing.T) {
	tests := []struct {
		partURI string
		expect  string
	}{
		{"word/document.xml", "word/_rels/document.xml.rels"},
		{"xl/workbook.xml", "xl/_rels/workbook.xml.rels"},
		{"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"},
		{"document.xml", "_rels/document.xml.rels"},
		{"word/header1.xml", "word/_rels/header1.xml.rels"},
	}

	for _, tt := range tests {
		got := RelationshipsPathForPart(tt.partURI)
		if got != tt.expect {
			t.Errorf("RelationshipsPathForPart(%q) = %q, want %q", tt.partURI, got, tt.expect)
		}
	}
}

func TestResolveRelationshipTarget(t *testing.T) {
	tests := []struct {
		source string
		target string
		expect string
	}{
		{"word/document.xml", "styles.xml", "word/styles.xml"},
		{"word/document.xml", "header1.xml", "word/header1.xml"},
		{"word/document.xml", "/word/styles.xml", "word/styles.xml"},
		{"xl/workbook.xml", "worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"},
		{"document.xml", "styles.xml", "styles.xml"},
	}

	for _, tt := range tests {
		got := ResolveRelationshipTarget(tt.source, tt.target)
		if got != tt.expect {
			t.Errorf("ResolveRelationshipTarget(%q, %q) = %q, want %q", tt.source, tt.target, got, tt.expect)
		}
	}
}

func TestRelationshipsAdd(t *testing.T) {
	rels := NewRelationships()

	rel1 := rels.Add(RelTypeStyles, "styles.xml", TargetModeInternal)
	if rel1.ID != "rId1" {
		t.Errorf("First relationship ID = %q, want rId1", rel1.ID)
	}

	rel2 := rels.Add(RelTypeSettings, "settings.xml", TargetModeInternal)
	if rel2.ID != "rId2" {
		t.Errorf("Second relationship ID = %q, want rId2", rel2.ID)
	}

	if len(rels.Relationships) != 2 {
		t.Errorf("len(Relationships) = %d, want 2", len(rels.Relationships))
	}
}

func TestRelationshipsByType(t *testing.T) {
	rels := NewRelationships()
	rels.Add(RelTypeStyles, "styles.xml", TargetModeInternal)
	rels.Add(RelTypeSettings, "settings.xml", TargetModeInternal)
	rels.Add(RelTypeHeader, "header1.xml", TargetModeInternal)
	rels.Add(RelTypeHeader, "header2.xml", TargetModeInternal)

	headers := rels.ByType(RelTypeHeader)
	if len(headers) != 2 {
		t.Errorf("ByType(Header) returned %d relationships, want 2", len(headers))
	}

	styles := rels.ByType(RelTypeStyles)
	if len(styles) != 1 {
		t.Errorf("ByType(Styles) returned %d relationships, want 1", len(styles))
	}
}

func TestRelationshipsFirstByType(t *testing.T) {
	rels := NewRelationships()
	rels.Add(RelTypeStyles, "styles.xml", TargetModeInternal)

	if rel := rels.FirstByType(RelTypeStyles); rel == nil {
		t.Error("FirstByType(Styles) returned nil")
	} else if rel.Target != "styles.xml" {
		t.Errorf("FirstByType(Styles).Target = %q, want styles.xml", rel.Target)
	}

	if rel := rels.FirstByType(RelTypeSettings); rel != nil {
		t.Error("FirstByType(Settings) should return nil")
	}
}

func TestRelationshipsRemove(t *testing.T) {
	rels := NewRelationships()
	rels.Add(RelTypeStyles, "styles.xml", TargetModeInternal)
	rels.Add(RelTypeSettings, "settings.xml", TargetModeInternal)

	if !rels.Remove("rId1") {
		t.Error("Remove(rId1) returned false")
	}
	if len(rels.Relationships) != 1 {
		t.Errorf("After remove, len = %d, want 1", len(rels.Relationships))
	}
	if rels.Remove("rId1") {
		t.Error("Second Remove(rId1) should return false")
	}
}
