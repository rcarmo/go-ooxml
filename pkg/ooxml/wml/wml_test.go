// Package wml tests for WordprocessingML types.
package wml

import (
	"encoding/xml"
	"strings"
	"testing"
)

// =============================================================================
// Document Type Tests
// =============================================================================

func TestDocument_Marshal(t *testing.T) {
	doc := &Document{
		Body: &Body{},
	}

	data, err := xml.Marshal(doc)
	if err != nil {
		t.Fatalf("Marshal error: %v", err)
	}

	s := string(data)
	if !strings.Contains(s, "document") {
		t.Error("marshaled output should contain 'document'")
	}
	if !strings.Contains(s, "body") {
		t.Error("marshaled output should contain 'body'")
	}
}

func TestDocument_Unmarshal(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`

	var doc Document
	err := xml.Unmarshal([]byte(xmlData), &doc)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if doc.Body == nil {
		t.Fatal("Body should not be nil")
	}

	if len(doc.Body.Content) != 1 {
		t.Fatalf("expected 1 element in body, got %d", len(doc.Body.Content))
	}

	p, ok := doc.Body.Content[0].(*P)
	if !ok {
		t.Fatalf("expected *P, got %T", doc.Body.Content[0])
	}

	if len(p.Content) != 1 {
		t.Fatalf("expected 1 run in paragraph, got %d", len(p.Content))
	}

	r, ok := p.Content[0].(*R)
	if !ok {
		t.Fatalf("expected *R, got %T", p.Content[0])
	}

	var text string
	for _, elem := range r.Content {
		if te, ok := elem.(*T); ok {
			text += te.Text
		}
	}
	if text != "Hello World" {
		t.Errorf("expected text 'Hello World', got %q", text)
	}
}

// =============================================================================
// Body Tests
// =============================================================================

func TestBody_Unmarshal_MultipleParagraphs(t *testing.T) {
	xmlData := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>First</w:t></w:r></w:p>
  <w:p><w:r><w:t>Second</w:t></w:r></w:p>
  <w:p><w:r><w:t>Third</w:t></w:r></w:p>
</w:body>`

	var body Body
	err := xml.Unmarshal([]byte(xmlData), &body)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if len(body.Content) != 3 {
		t.Errorf("expected 3 paragraphs, got %d", len(body.Content))
	}
}

func TestBody_Unmarshal_WithTable(t *testing.T) {
	xmlData := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Before table</w:t></w:r></w:p>
  <w:tbl>
    <w:tr>
      <w:tc><w:p><w:r><w:t>Cell1</w:t></w:r></w:p></w:tc>
    </w:tr>
  </w:tbl>
  <w:p><w:r><w:t>After table</w:t></w:r></w:p>
</w:body>`

	var body Body
	err := xml.Unmarshal([]byte(xmlData), &body)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if len(body.Content) != 3 {
		t.Fatalf("expected 3 elements (p, tbl, p), got %d", len(body.Content))
	}

	// Check types
	if _, ok := body.Content[0].(*P); !ok {
		t.Errorf("element 0: expected *P, got %T", body.Content[0])
	}
	if _, ok := body.Content[1].(*Tbl); !ok {
		t.Errorf("element 1: expected *Tbl, got %T", body.Content[1])
	}
	if _, ok := body.Content[2].(*P); !ok {
		t.Errorf("element 2: expected *P, got %T", body.Content[2])
	}
}

func TestBody_Unmarshal_WithSectPr(t *testing.T) {
	xmlData := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Content</w:t></w:r></w:p>
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
</w:body>`

	var body Body
	err := xml.Unmarshal([]byte(xmlData), &body)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if body.SectPr == nil {
		t.Error("SectPr should not be nil")
	}
}

// =============================================================================
// Paragraph Tests
// =============================================================================

func TestParagraph_Unmarshal_WithStyle(t *testing.T) {
	xmlData := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r><w:t>Title</w:t></w:r>
</w:p>`

	var p P
	err := xml.Unmarshal([]byte(xmlData), &p)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if p.PPr == nil {
		t.Fatal("PPr should not be nil")
	}
	if p.PPr.PStyle == nil || p.PPr.PStyle.Val != "Heading1" {
		t.Error("style should be Heading1")
	}
	if p.PPr.Jc == nil || p.PPr.Jc.Val != "center" {
		t.Error("alignment should be center")
	}
}

func TestParagraph_Unmarshal_WithInsertions(t *testing.T) {
	xmlData := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:r><w:t>Original </w:t></w:r>
  <w:ins w:id="1" w:author="Author" w:date="2024-01-01T00:00:00Z">
    <w:r><w:t>inserted text</w:t></w:r>
  </w:ins>
</w:p>`

	var p P
	err := xml.Unmarshal([]byte(xmlData), &p)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if len(p.Content) != 2 {
		t.Fatalf("expected 2 elements (r, ins), got %d", len(p.Content))
	}

	ins, ok := p.Content[1].(*Ins)
	if !ok {
		t.Fatalf("expected *Ins, got %T", p.Content[1])
	}
	if ins.Author != "Author" {
		t.Errorf("author = %q, want 'Author'", ins.Author)
	}
}

func TestParagraph_Unmarshal_WithDeletions(t *testing.T) {
	xmlData := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:del w:id="2" w:author="Editor" w:date="2024-01-02T00:00:00Z">
    <w:r><w:delText>deleted text</w:delText></w:r>
  </w:del>
  <w:r><w:t>remaining</w:t></w:r>
</w:p>`

	var p P
	err := xml.Unmarshal([]byte(xmlData), &p)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if len(p.Content) != 2 {
		t.Fatalf("expected 2 elements, got %d", len(p.Content))
	}

	del, ok := p.Content[0].(*Del)
	if !ok {
		t.Fatalf("expected *Del, got %T", p.Content[0])
	}
	if del.Author != "Editor" {
		t.Errorf("author = %q, want 'Editor'", del.Author)
	}
}

// =============================================================================
// Run Tests
// =============================================================================

func TestRun_Unmarshal_WithFormatting(t *testing.T) {
	xmlData := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr>
    <w:b/>
    <w:i/>
    <w:u w:val="single"/>
    <w:sz w:val="28"/>
    <w:color w:val="FF0000"/>
  </w:rPr>
  <w:t>Formatted Text</w:t>
</w:r>`

	var r R
	err := xml.Unmarshal([]byte(xmlData), &r)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if r.RPr == nil {
		t.Fatal("RPr should not be nil")
	}
	if r.RPr.B == nil {
		t.Error("bold should be set")
	}
	if r.RPr.I == nil {
		t.Error("italic should be set")
	}
	if r.RPr.U == nil || r.RPr.U.Val != "single" {
		t.Error("underline should be single")
	}
	if r.RPr.Sz == nil || r.RPr.Sz.Val != 28 {
		t.Errorf("font size = %v, want 28", r.RPr.Sz)
	}
	if r.RPr.Color == nil || r.RPr.Color.Val != "FF0000" {
		t.Error("color should be FF0000")
	}

	// Check text content
	found := false
	for _, elem := range r.Content {
		if te, ok := elem.(*T); ok {
			if te.Text == "Formatted Text" {
				found = true
			}
		}
	}
	if !found {
		t.Error("text content not found")
	}
}

func TestRun_Unmarshal_WithBreaks(t *testing.T) {
	xmlData := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:t>Line 1</w:t>
  <w:br/>
  <w:t>Line 2</w:t>
  <w:tab/>
  <w:t>Tabbed</w:t>
</w:r>`

	var r R
	err := xml.Unmarshal([]byte(xmlData), &r)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	// Should have 5 content elements
	if len(r.Content) != 5 {
		t.Errorf("expected 5 content elements, got %d", len(r.Content))
	}
}

// =============================================================================
// Table Tests
// =============================================================================

func TestTable_Unmarshal(t *testing.T) {
	xmlData := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="2500"/>
    <w:gridCol w:w="2500"/>
  </w:tblGrid>
  <w:tr>
    <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>`

	var tbl Tbl
	err := xml.Unmarshal([]byte(xmlData), &tbl)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if len(tbl.Tr) != 2 {
		t.Errorf("expected 2 rows, got %d", len(tbl.Tr))
	}

	if len(tbl.Tr[0].Tc) != 2 {
		t.Errorf("expected 2 cells in first row, got %d", len(tbl.Tr[0].Tc))
	}

	// Check first cell content
	tc := tbl.Tr[0].Tc[0]
	if len(tc.Content) != 1 {
		t.Fatalf("expected 1 paragraph in cell, got %d", len(tc.Content))
	}

	p, ok := tc.Content[0].(*P)
	if !ok {
		t.Fatalf("expected *P, got %T", tc.Content[0])
	}
	if len(p.Content) != 1 {
		t.Fatalf("expected 1 run in paragraph, got %d", len(p.Content))
	}
}

func TestTable_Marshal_RoundTrip(t *testing.T) {
	// Create a table
	tbl := &Tbl{
		TblPr: &TblPr{
			TblStyle: &TblStyle{Val: "TableGrid"},
		},
		TblGrid: &TblGrid{
			GridCol: []*GridCol{{W: 2500}, {W: 2500}},
		},
		Tr: []*Tr{
			{
				Tc: []*Tc{
					{Content: []interface{}{&P{Content: []interface{}{&R{Content: []interface{}{&T{Text: "A1"}}}}}}},
					{Content: []interface{}{&P{Content: []interface{}{&R{Content: []interface{}{&T{Text: "B1"}}}}}}},
				},
			},
			{
				Tc: []*Tc{
					{Content: []interface{}{&P{Content: []interface{}{&R{Content: []interface{}{&T{Text: "A2"}}}}}}},
					{Content: []interface{}{&P{Content: []interface{}{&R{Content: []interface{}{&T{Text: "B2"}}}}}}},
				},
			},
		},
	}

	// Marshal
	data, err := xml.Marshal(tbl)
	if err != nil {
		t.Fatalf("Marshal error: %v", err)
	}

	// Unmarshal into new struct
	var tbl2 Tbl
	err = xml.Unmarshal(data, &tbl2)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	// Verify
	if len(tbl2.Tr) != 2 {
		t.Errorf("expected 2 rows after round-trip, got %d", len(tbl2.Tr))
	}
	if len(tbl2.Tr[0].Tc) != 2 {
		t.Errorf("expected 2 cells after round-trip, got %d", len(tbl2.Tr[0].Tc))
	}
}

// =============================================================================
// Cell Tests
// =============================================================================

func TestCell_Unmarshal_WithProperties(t *testing.T) {
	xmlData := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tcPr>
    <w:tcW w:w="2880" w:type="dxa"/>
    <w:gridSpan w:val="2"/>
    <w:vMerge w:val="restart"/>
    <w:shd w:val="clear" w:fill="FFFF00"/>
  </w:tcPr>
  <w:p><w:r><w:t>Merged Cell</w:t></w:r></w:p>
</w:tc>`

	var tc Tc
	err := xml.Unmarshal([]byte(xmlData), &tc)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	if tc.TcPr == nil {
		t.Fatal("TcPr should not be nil")
	}
	if tc.TcPr.GridSpan == nil || tc.TcPr.GridSpan.Val != 2 {
		t.Error("grid span should be 2")
	}
	if tc.TcPr.VMerge == nil || tc.TcPr.VMerge.Val != "restart" {
		t.Error("vMerge should be restart")
	}
	if tc.TcPr.Shd == nil || tc.TcPr.Shd.Fill != "FFFF00" {
		t.Error("shading fill should be FFFF00")
	}
}

// =============================================================================
// OnOff Type Tests
// =============================================================================

func TestOnOff_Enabled(t *testing.T) {
	trueVal := true
	falseVal := false

	tests := []struct {
		name    string
		onoff   *OnOff
		enabled bool
	}{
		{"nil", nil, false},
		{"empty (presence=true)", &OnOff{}, true},
		{"val=true", &OnOff{Val: &trueVal}, true},
		{"val=false", &OnOff{Val: &falseVal}, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := tt.onoff.Enabled()
			if got != tt.enabled {
				t.Errorf("Enabled() = %v, want %v", got, tt.enabled)
			}
		})
	}
}

func TestNewOnOffEnabled(t *testing.T) {
	onoff := NewOnOffEnabled()
	if onoff == nil {
		t.Fatal("NewOnOffEnabled() returned nil")
	}
	if !onoff.Enabled() {
		t.Error("new OnOff should be enabled")
	}
}

func TestNewOnOff(t *testing.T) {
	enabled := NewOnOff(true)
	if !enabled.Enabled() {
		t.Error("NewOnOff(true) should be enabled")
	}

	disabled := NewOnOff(false)
	if disabled.Enabled() {
		t.Error("NewOnOff(false) should not be enabled")
	}
}

// =============================================================================
// Run Properties Tests
// =============================================================================

func TestRPr_Marshal(t *testing.T) {
	rPr := &RPr{
		B:         &OnOff{},
		I:         &OnOff{},
		U:         &U{Val: "single"},
		Strike:    &OnOff{},
		Sz:        &Sz{Val: 24},
		SzCs:      &Sz{Val: 24},
		Color:     &Color{Val: "FF0000"},
		Highlight: &Highlight{Val: "yellow"},
		RFonts:    &RFonts{Ascii: "Arial"},
	}

	data, err := xml.Marshal(rPr)
	if err != nil {
		t.Fatalf("Marshal error: %v", err)
	}

	s := string(data)
	if !strings.Contains(s, "b") {
		t.Error("should contain bold element")
	}
	if !strings.Contains(s, "i") {
		t.Error("should contain italic element")
	}
	if !strings.Contains(s, "u") {
		t.Error("should contain underline element")
	}
}

// =============================================================================
// Paragraph Properties Tests
// =============================================================================

func TestPPr_Marshal(t *testing.T) {
	before := int64(240)
	after := int64(120)
	line := int64(276)

	pPr := &PPr{
		PStyle: &PStyle{Val: "Heading1"},
		Jc:     &Jc{Val: "center"},
		Spacing: &Spacing{
			Before: &before,
			After:  &after,
			Line:   &line,
		},
		KeepNext: &OnOff{},
	}

	data, err := xml.Marshal(pPr)
	if err != nil {
		t.Fatalf("Marshal error: %v", err)
	}

	s := string(data)
	if !strings.Contains(s, "pStyle") {
		t.Error("should contain pStyle element")
	}
	if !strings.Contains(s, "jc") {
		t.Error("should contain jc element")
	}
	if !strings.Contains(s, "spacing") {
		t.Error("should contain spacing element")
	}
}

// =============================================================================
// Full Document Round-Trip Test
// =============================================================================

func TestDocument_RoundTrip(t *testing.T) {
	// Create document
	doc := &Document{
		Body: &Body{
			Content: []interface{}{
				&P{
					PPr: &PPr{
						PStyle: &PStyle{Val: "Heading1"},
					},
					Content: []interface{}{
						&R{
							RPr: &RPr{B: &OnOff{}},
							Content: []interface{}{
								&T{Text: "Title"},
							},
						},
					},
				},
				&P{
					Content: []interface{}{
						&R{
							Content: []interface{}{
								&T{Text: "Body paragraph"},
							},
						},
					},
				},
				&Tbl{
					Tr: []*Tr{
						{
							Tc: []*Tc{
								{Content: []interface{}{&P{Content: []interface{}{&R{Content: []interface{}{&T{Text: "Cell"}}}}}}},
							},
						},
					},
				},
			},
			SectPr: &SectPr{
				PgSz: &PgSz{W: 12240, H: 15840},
			},
		},
	}

	// Marshal
	data, err := xml.Marshal(doc)
	if err != nil {
		t.Fatalf("Marshal error: %v", err)
	}

	// Unmarshal
	var doc2 Document
	err = xml.Unmarshal(data, &doc2)
	if err != nil {
		t.Fatalf("Unmarshal error: %v", err)
	}

	// Verify structure
	if doc2.Body == nil {
		t.Fatal("Body should not be nil")
	}
	if len(doc2.Body.Content) != 3 {
		t.Errorf("expected 3 content elements, got %d", len(doc2.Body.Content))
	}

	// Check first paragraph
	p, ok := doc2.Body.Content[0].(*P)
	if !ok {
		t.Fatalf("first element should be *P, got %T", doc2.Body.Content[0])
	}
	if p.PPr == nil || p.PPr.PStyle == nil || p.PPr.PStyle.Val != "Heading1" {
		t.Error("first paragraph style should be Heading1")
	}

	// Check table
	tbl, ok := doc2.Body.Content[2].(*Tbl)
	if !ok {
		t.Fatalf("third element should be *Tbl, got %T", doc2.Body.Content[2])
	}
	if len(tbl.Tr) != 1 {
		t.Error("table should have 1 row")
	}
}

// =============================================================================
// Text Element Tests
// =============================================================================

func TestNewT(t *testing.T) {
	tests := []struct {
		name           string
		text           string
		expectPreserve bool
	}{
		{"simple", "hello", false},
		{"leading space", " hello", true},
		{"trailing space", "hello ", true},
		{"double space", "hello  world", true},
		{"no space", "helloworld", false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			te := NewT(tt.text)
			if te.Text != tt.text {
				t.Errorf("Text = %q, want %q", te.Text, tt.text)
			}
			hasPreserve := te.Space == "preserve"
			if hasPreserve != tt.expectPreserve {
				t.Errorf("Space = %q, expectPreserve = %v", te.Space, tt.expectPreserve)
			}
		})
	}
}

