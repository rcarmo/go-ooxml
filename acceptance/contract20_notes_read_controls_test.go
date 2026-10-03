package acceptance

import (
	"bytes"
	"errors"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Test-only XML projection controls: different valid source topologies make
// breaks, empty paragraphs and literal field values independently observable.
func TestContract20NotesOracleAlternates(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	const head = `<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>`
	const tail = `</p:spTree></p:cSld></p:notes>`
	for _, tc := range []struct{ name, body, want string }{
		{"blank-interior", `<a:p><a:r><a:t>A</a:t></a:r></a:p><a:p/><a:p><a:r><a:t>B</a:t></a:r></a:p>`, "A\n\nB"},
		{"break-and-field", `<a:p><a:r><a:t>First</a:t></a:r><a:br/><a:fld type="datetimeFigureOut"><a:t>Stored &amp; visible</a:t></a:fld></a:p>`, "First\nStored & visible"},
		{"edge-blank", `<a:p/><a:p><a:r><a:t>Middle</a:t></a:r></a:p><a:p/>`, "\nMiddle\n"},
		{"boundary-crlf", `<a:p><a:r><a:t>A&#13;</a:t></a:r><a:r><a:t>&#10;B</a:t></a:r></a:p>`, "A\n\nB"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			xml := []byte(head + `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody>` + tc.body + `</p:txBody></p:sp>` + tail)
			original := bytes.Clone(xml)
			got, err := contract20OraclePlaceholderText(xml, "\n", "body")
			if err != nil || got != tc.want {
				t.Fatalf("oracle %q, expected %q: %v", got, tc.want, err)
			}
			if !bytes.Equal(xml, original) {
				t.Fatal("oracle altered caller input")
			}
		})
	}
}

// The admitted formatting subset is intentionally bounded. These known
// property elements may be present but cannot hide projected notes content.
func TestContract20NotesReadFormattingPositive(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	w := &contract20NotesReadWorld{}
	if err := w.input(); err != nil {
		t.Fatal(err)
	}
	parts := make(map[string][]byte, len(w.members["ordered-notes"]))
	for name, payload := range w.members["ordered-notes"] {
		parts[name] = bytes.Clone(payload)
	}
	const notes = "ppt/notesSlides/notesSlide5.xml"
	const old = `<a:r><a:t>Key themes:</a:t></a:r>`
	const formatted = `<a:r><a:rPr b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="F&amp;F"/></a:rPr><a:t>Key themes:</a:t></a:r>`
	if strings.Count(string(parts[notes]), old) != 1 {
		t.Fatal("sealed run anchor ambiguous")
	}
	parts[notes] = []byte(strings.Replace(string(parts[notes]), old, formatted, 1))
	archive, err := contract20PackMembers(parts)
	if err != nil {
		t.Fatal(err)
	}
	original := bytes.Clone(archive)
	before, err := contract20TableGraph(archive)
	if err != nil {
		t.Fatal(err)
	}
	got, err := presentation.ReadContractSlideNotes(archive)
	if err != nil || len(got) != 5 {
		t.Fatalf("valid formatting notes: %v / %v", got, err)
	}
	const expected = "Key themes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12"
	oracle, err := contract20OraclePlaceholderText(parts[notes], "\n", "body")
	if err != nil || got[0].Text != expected || oracle != expected {
		t.Fatalf("valid formatting: reader %q oracle %q expected %q: %v", got[0].Text, oracle, expected, err)
	}
	afterParts, err := blankMemberPayloads(archive)
	if err != nil {
		t.Fatal(err)
	}
	after, err := contract20TableGraph(archive)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(archive, original) || !reflect.DeepEqual(afterParts, parts) || !reflect.DeepEqual(after, before) {
		t.Fatal("read changed source archive, member payloads or graph")
	}
}

// A literal XML character reference to CR is a stored visible character, not
// a paragraph break. Both the production reader and independent XML decoder
// must normalise it without evaluating the field or changing any source part.
func TestContract20NotesReadCRAlternate(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	w := &contract20NotesReadWorld{}
	if err := w.input(); err != nil {
		t.Fatal(err)
	}
	parts := make(map[string][]byte, len(w.members["ordered-notes"]))
	for name, payload := range w.members["ordered-notes"] {
		parts[name] = bytes.Clone(payload)
	}
	const notes = "ppt/notesSlides/notesSlide5.xml"
	if strings.Count(string(parts[notes]), `<a:t>Key themes:</a:t>`) != 1 {
		t.Fatal("sealed visible text anchor ambiguous")
	}
	parts[notes] = []byte(strings.Replace(string(parts[notes]), `<a:t>Key themes:</a:t>`, `<a:t>Key&#13;themes:</a:t>`, 1))
	archive, err := contract20PackMembers(parts)
	if err != nil {
		t.Fatal(err)
	}
	original := bytes.Clone(archive)
	before, err := contract20TableGraph(archive)
	if err != nil {
		t.Fatal(err)
	}
	got, err := presentation.ReadContractSlideNotes(archive)
	if err != nil || len(got) != 5 {
		t.Fatalf("CR alternate notes: %v / %v", got, err)
	}
	const expected = "Key\nthemes:\ncreation, responsibility, isolation\n\nVisible date: 2026-03-12"
	oracle, err := contract20OraclePlaceholderText(parts[notes], "\n", "body")
	if err != nil || got[0].Text != expected || oracle != expected {
		t.Fatalf("CR alternate: reader %q, oracle %q, expected %q: %v", got[0].Text, oracle, expected, err)
	}
	afterParts, err := blankMemberPayloads(archive)
	if err != nil {
		t.Fatal(err)
	}
	after, err := contract20TableGraph(archive)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(archive, original) || !reflect.DeepEqual(afterParts, parts) || !reflect.DeepEqual(after, before) {
		t.Fatal("read changed source archive, member payloads or graph")
	}
}

// These are native negative controls, not new canonical Contract20 cases.
func TestContract20NotesReadCollateralBatch(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	w := &contract20NotesReadWorld{}
	if err := w.input(); err != nil {
		t.Fatal(err)
	}
	original := bytes.Clone(w.sources["ordered-notes"])
	base := w.members["ordered-notes"]
	baseline := make(map[string][]byte, len(base))
	for name, payload := range base {
		baseline[name] = bytes.Clone(payload)
	}
	baselineGraph := w.graphs["ordered-notes"]
	for _, tc := range []struct {
		name, part, before, after, kind string
	}{
		{"duplicate-slide-id", "ppt/presentation.xml", `r:id="rId12"`, `r:id="rId7"`, "PPTX_PRESENTATION_INVALID"},
		{"duplicate-notes-body", "ppt/notesSlides/notesSlide5.xml", `<p:ph type="body" idx="3" sz="quarter"/>`, `<p:ph type="body" idx="3" sz="quarter"/><p:ph type="body" idx="4"/>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"foreign-notes-child", "ppt/notesSlides/notesSlide5.xml", `<a:br/>`, `<a:tab/>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"nested-foreign-run-child", "ppt/notesSlides/notesSlide5.xml", `<a:r><a:t>Key themes:</a:t></a:r>`, `<a:r><a:tab/><a:t>Key themes:</a:t></a:r>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"nested-foreign-field-child", "ppt/notesSlides/notesSlide5.xml", `<a:t>2026-03-12</a:t></a:fld>`, `<a:tab/><a:t>2026-03-12</a:t></a:fld>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"foreign-body-child", "ppt/notesSlides/notesSlide5.xml", `<a:lstStyle/>`, `<a:lstStyle/><a:tab/>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"visible-text-in-run-property", "ppt/notesSlides/notesSlide5.xml", `<a:r><a:t>Key themes:</a:t></a:r>`, `<a:r><a:rPr><a:t>LOST VISIBLE</a:t></a:rPr><a:t>Key themes:</a:t></a:r>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"foreign-run-property", "ppt/notesSlides/notesSlide5.xml", `<a:r><a:t>Key themes:</a:t></a:r>`, `<a:r><a:rPr><z:unmodelled xmlns:z="urn:foreign"/></a:rPr><a:t>Key themes:</a:t></a:r>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"break-in-paragraph-property", "ppt/notesSlides/notesSlide5.xml", `<a:p><a:r><a:t>Key themes:</a:t>`, `<a:p><a:pPr><a:br/></a:pPr><a:r><a:t>Key themes:</a:t>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"foreign-body-property", "ppt/notesSlides/notesSlide5.xml", `<a:lstStyle/>`, `<a:lstStyle><z:unmodelled xmlns:z="urn:foreign"/></a:lstStyle>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"nested-unknown-owner", "ppt/notesSlides/notesSlide5.xml", `<a:r><a:t>Key themes:</a:t></a:r>`, `<a:r><a:rPr><a:solidFill><a:t>hidden</a:t></a:solidFill></a:rPr><a:t>Key themes:</a:t></a:r>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"wrong-notes-content-type", "[Content_Types].xml", `PartName="/ppt/notesSlides/notesSlide5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"`, `PartName="/ppt/notesSlides/notesSlide5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
		{"shared-notes-owner", "ppt/slides/_rels/slide4.xml.rels", `</Relationships>`, `<Relationship Id="contract20Duplicate" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide5.xml"/></Relationships>`, "PPTX_NOTES_STRUCTURE_UNSUPPORTED"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := make(map[string][]byte, len(base))
			for name, payload := range base {
				parts[name] = bytes.Clone(payload)
			}
			if strings.Count(string(parts[tc.part]), tc.before) != 1 {
				t.Fatalf("control source anchor not unique: %s", tc.before)
			}
			parts[tc.part] = []byte(strings.Replace(string(parts[tc.part]), tc.before, tc.after, 1))
			archive, err := contract20PackMembers(parts)
			if err != nil {
				t.Fatal(err)
			}
			variant := bytes.Clone(archive)
			variantMembers, err := blankMemberPayloads(archive)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(variantMembers, parts) {
				t.Fatal("control packing changed source members")
			}
			variantGraph, err := contract20TableGraph(archive)
			if err != nil {
				t.Fatal(err)
			}
			got, err := presentation.ReadContractSlideNotes(archive)
			var refused *packaging.Refusal
			if !errors.As(err, &refused) || refused.Kind != tc.kind || got != nil {
				t.Fatalf("variant %s returned %v/%v", tc.name, got, err)
			}
			afterMembers, err := blankMemberPayloads(archive)
			if err != nil {
				t.Fatal(err)
			}
			afterGraph, err := contract20TableGraph(archive)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(archive, variant) || !reflect.DeepEqual(afterMembers, variantMembers) || !reflect.DeepEqual(afterGraph, variantGraph) {
				t.Fatal("refusal modified variant archive, member payloads or graph")
			}
			originalMembers, err := blankMemberPayloads(w.sources["ordered-notes"])
			if err != nil {
				t.Fatal(err)
			}
			originalGraph, err := contract20TableGraph(w.sources["ordered-notes"])
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(w.sources["ordered-notes"], original) || !reflect.DeepEqual(originalMembers, baseline) || !reflect.DeepEqual(originalGraph, baselineGraph) {
				t.Fatal("control modified original archive, member payloads or graph")
			}
		})
	}
}
