package acceptance

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"sort"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Owned notes inputs and a native synthetic paragraph variant retain the same
// no-op, edit, clear, reopen, payload-budget and placeholder assertions.
func TestOwnedNotesFixtures(t *testing.T) {
	root := testutil.FixturePath()
	observations := []map[string]any{}
	for _, f := range []struct {
		name, path, sha string
		variant         bool
	}{
		{"owned notes", "pptx/notes.pptx", "04faba67841dda25dc3ff9e3e6e345e6feeeef1cf25a6b9065bf5fbdc83163dc", false},
		{"native blank paragraph variant", "pptx/notes.pptx", "04faba67841dda25dc3ff9e3e6e345e6feeeef1cf25a6b9065bf5fbdc83163dc", true},
	} {
		t.Run(f.name, func(t *testing.T) {
			source, err := pinnedFile(root, f.path, f.sha)
			if err != nil {
				t.Fatal(err)
			}
			original := bytes.Clone(source)
			wantInitial := "Remember to emphasize the Gothic elements"
			if f.variant {
				parts, err := zipPayloads(source)
				if err != nil {
					t.Fatal(err)
				}
				name := "ppt/notesSlides/notesSlide1.xml"
				parts[name] = bytes.Replace(parts[name], []byte(`</p:txBody>`), []byte(`<a:p/></p:txBody>`), 1)
				var output bytes.Buffer
				zw := zip.NewWriter(&output)
				names := make([]string, 0, len(parts))
				for name := range parts {
					names = append(names, name)
				}
				sort.Strings(names)
				for _, name := range names {
					w, err := zw.Create(name)
					if err != nil {
						t.Fatal(err)
					}
					if _, err = w.Write(parts[name]); err != nil {
						t.Fatal(err)
					}
				}
				if err := zw.Close(); err != nil {
					t.Fatal(err)
				}
				source = output.Bytes()
				wantInitial += "\n"
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
			if n.Text() != wantInitial {
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
			if err != nil || !bytes.Equal(original, sourceNow) {
				t.Fatal("source fixture changed", err)
			}
			observations = append(observations, map[string]any{"name": f.name, "path": f.path, "sha256": f.sha, "syntheticVariant": f.variant, "inputSHA256": sha256hex(source), "initialText": n.Text(), "multilineReadback": got.Text(), "emptyReadback": emptyNotes.Text(), "changedParts": receipt.Changes, "sourceUnchanged": true, "noOpByteIdentical": true, "otherPayloadsExact": true, "otherPlaceholdersExact": true, "otherPlaceholderCount": len(unchangedShapes), "outputSHA256": sha256hex(output)})
		})
	}
	dir := os.Getenv("OOXML_REPORT_DIR")
	if dir == "" {
		dir = "../reports/acceptance"
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(dir, "owned-notes-readback.json"), map[string]any{"schema": 1, "referenceDistribution": loadReferencePin(t).Commit, "subject": "native-library", "transport": "none", "observations": observations, "pythonExecuted": false, "officeExecuted": false})
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
