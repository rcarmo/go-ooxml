package acceptance

import (
	"bytes"
	"encoding/xml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Explicitly requested read-only source fixture checks; no Python runtime or
// fixture redistribution. Missing/hash-mismatched inputs fail requested runs.
func TestRetiredNotesFixtures(t *testing.T) {
	root := testutil.ReferencePath("reference-assets", "pptx")
	observations := []map[string]any{}
	for _, f := range []struct{ path, sha string }{{"[retired external test identity]", "72f375efdbecdb8b95adf37c8fc753ec5416b4353b48bad8f6b89c0dedb74abf"}, {"[retired external test identity]", "a4ce1dae2558ced03ea5df2a415c32b57c9a10a0eedc8a1bfd60e588ee7cbf98"}} {
		t.Run(f.path, func(t *testing.T) {
			source, err := pinnedFile(root, f.path, f.sha)
			if err != nil {
				t.Fatal(err)
			}
			s, err := presentation.OpenEditing(source, packaging.Limits{MaxSourceBytes: 32 << 20, MaxEntries: 4096, MaxPartBytes: 16 << 20, MaxTotalBytes: 64 << 20})
			if err != nil {
				t.Fatal(err)
			}
			const slide = "ppt/slides/slide1.xml"
			const part = "ppt/notesSlides/notesSlide1.xml"
			n, err := s.FindNotes(slide)
			if err != nil {
				t.Fatal(err)
			}
			if n.Text() != "Speaker notes for the clone fixture." {
				t.Fatalf("initial readback %q", n.Text())
			}
			dir := t.TempDir()
			same := filepath.Join(dir, "noop.pptx")
			if err = s.ReplaceNotes(n, n.Text()); err != nil {
				t.Fatal(err)
			}
			if _, err = s.SaveAs(same); err != nil {
				t.Fatal(err)
			}
			noOp, err := os.ReadFile(same)
			if err != nil || !bytes.Equal(source, noOp) {
				t.Fatal("no-op source bytes changed", err)
			}
			const text = "First line\nSecond line\n\nFourth line"
			if err = s.ReplaceNotes(n, text); err != nil {
				t.Fatal(err)
			}
			dest := filepath.Join(dir, "changed.pptx")
			receipt, err := s.SaveAs(dest)
			if err != nil {
				t.Fatal(err)
			}
			if len(receipt.Changes) != 1 || receipt.Changes[0].Part != part {
				t.Fatal("changed-part budget", receipt)
			}
			output, err := os.ReadFile(dest)
			if err != nil {
				t.Fatal(err)
			}
			before, _ := zipPayloads(source)
			after, err := zipPayloads(output)
			if err != nil || len(before) != len(after) {
				t.Fatal("part inventory", err)
			}
			for name, data := range before {
				if name != part && !bytes.Equal(data, after[name]) {
					t.Fatal("unrelated payload changed", name)
				}
			}
			unchangedShapes := otherNotesShapes(t, before[part])
			afterShapes := otherNotesShapes(t, after[part])
			if len(unchangedShapes) != len(afterShapes) {
				t.Fatal("other notes shapes inventory changed")
			}
			for i, b := range unchangedShapes {
				if !bytes.Equal(b, afterShapes[i]) {
					t.Fatal("other notes placeholder bytes changed")
				}
			}
			reopened, err := presentation.OpenEditing(output, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			got, err := reopened.FindNotes(slide)
			if err != nil || got.Text() != text {
				t.Fatal("multiline reopen", err)
			}
			if err = reopened.ReplaceNotes(got, ""); err != nil {
				t.Fatal(err)
			}
			empty := filepath.Join(dir, "empty.pptx")
			if _, err = reopened.SaveAs(empty); err != nil {
				t.Fatal(err)
			}
			emptyBytes, err := os.ReadFile(empty)
			if err != nil {
				t.Fatal(err)
			}
			again, err := presentation.OpenEditing(emptyBytes, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			emptyNotes, err := again.FindNotes(slide)
			if err != nil || emptyNotes.Text() != "" {
				t.Fatal("empty reopen", err)
			}
			sourceNow, err := os.ReadFile(filepath.Join(root, f.path))
			if err != nil || !bytes.Equal(source, sourceNow) {
				t.Fatal("source fixture changed", err)
			}
			observations = append(observations, map[string]any{"path": f.path, "sha256": f.sha, "initialText": n.Text(), "multilineReadback": got.Text(), "emptyReadback": emptyNotes.Text(), "changedParts": receipt.Changes, "sourceUnchanged": true, "noOpByteIdentical": true, "otherPayloadsExact": true, "otherPlaceholdersExact": true, "otherPlaceholderCount": len(unchangedShapes), "outputSHA256": sha256hex(output)})
		})
	}
	dir := os.Getenv("OOXML_REPORT_DIR")
	if dir == "" {
		dir = "../reports/acceptance"
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(dir, "retired-notes-readback.json"), map[string]any{"schema": 1, "sourceRevision": "[retired implementation revision]", "subject": "native-library", "transport": "none", "observations": observations, "pythonExecuted": false, "officeExecuted": false})
}

func otherNotesShapes(t *testing.T, data []byte) [][]byte {
	t.Helper()
	doc, err := losslessxml.Parse(data)
	if err != nil {
		t.Fatal(err)
	}
	var result [][]byte
	for _, shape := range doc.Elements() {
		if shape.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "sp"}) {
			continue
		}
		body := false
		for _, e := range doc.Elements() {
			if e.Name() != (xml.Name{Space: packaging.NSPresentationML, Local: "ph"}) {
				continue
			}
			inside := false
			for p, ok := e.Parent(); ok; p, ok = p.Parent() {
				if p == shape {
					inside = true
					break
				}
			}
			if !inside {
				continue
			}
			for _, a := range e.Attributes() {
				if a.Name == (xml.Name{Local: "type"}) && a.Value == "body" {
					body = true
				}
			}
		}
		if !body {
			result = append(result, shape.Raw())
		}
	}
	return result
}
