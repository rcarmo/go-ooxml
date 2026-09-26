package presentation

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func openNotesFixture(t *testing.T) (*EditSession, []byte) {
	t.Helper()
	data, err := os.ReadFile("../../testdata/pptx/notes.pptx")
	if err != nil {
		t.Fatal(err)
	}
	hash := sha256.Sum256(data)
	if hex.EncodeToString(hash[:]) != "04faba67841dda25dc3ff9e3e6e345e6feeeef1cf25a6b9065bf5fbdc83163dc" {
		t.Fatal("notes source fixture changed")
	}
	s, err := OpenEditing(data, packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	return s, data
}
func TestExistingNotesFixtureBatch(t *testing.T) {
	const slide = "ppt/slides/slide1.xml"
	const part = "ppt/notesSlides/notesSlide1.xml"
	t.Run("exact part budget and existing placeholders", func(t *testing.T) {
		s, source := openNotesFixture(t)
		target, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		if target.Text() != "Remember to emphasize the Gothic elements" {
			t.Fatalf("unexpected notes %q", target.Text())
		}
		text := "Updated speaker notes"
		before, _, err := s.pkg.Part(part)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(target, text); err != nil {
			t.Fatal(err)
		}
		after, _, _ := s.pkg.Part(part)
		want := bytes.Replace(before, []byte(target.Text()), []byte(text), 1)
		if !bytes.Equal(want, after) {
			t.Fatal("notes XML changed beyond text")
		}
		dest := filepath.Join(t.TempDir(), "out.pptx")
		receipt, err := s.SaveAs(dest)
		if err != nil {
			t.Fatal(err)
		}
		if len(receipt.Changes) != 1 || receipt.Changes[0].Part != part {
			t.Fatal(receipt)
		}
		data, err := os.ReadFile(dest)
		if err != nil {
			t.Fatal(err)
		}
		q, err := packaging.OpenPreserved(source, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		g, err := q.Graph()
		if err != nil {
			t.Fatal(err)
		}
		delivered, err := OpenEditing(data, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		for _, p := range g.Parts {
			if p.Name == part {
				continue
			}
			a, _, _ := q.Part(p.Name)
			b, _, err := delivered.pkg.Part(p.Name)
			if err != nil || !bytes.Equal(a, b) {
				t.Fatal("unrelated payload changed", p.Name, err)
			}
		}
		r, err := delivered.FindNotes(slide)
		if err != nil || r.Text() != text {
			t.Fatal(r, err)
		}
	})
	t.Run("no-op rollback foreign stale handles and empty output", func(t *testing.T) {
		s, source := openNotesFixture(t)
		target, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		other, _ := openNotesFixture(t)
		if err = other.ReplaceNotes(target, "foreign"); err == nil {
			t.Fatal("foreign target accepted")
		}
		for _, text := range []string{"bad\x00", string([]byte{0xff}), "two\tcolumns"} {
			if err = s.ReplaceNotes(target, text); err == nil {
				t.Fatal("invalid accepted")
			}
		}
		if err = s.ReplaceNotes(target, target.Text()); err != nil {
			t.Fatal(err)
		}
		var b bytes.Buffer
		if err = s.pkg.WriteTo(&b); err != nil || !bytes.Equal(source, b.Bytes()) {
			t.Fatal("no-op/refusal changed archive")
		}
		second, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(target, ""); err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(second, "stale"); err == nil {
			t.Fatal("held target not stale")
		}
		fresh, err := s.FindNotes(slide)
		if err != nil || fresh.Text() != "" {
			t.Fatal("empty body failed", err)
		}
		if err = s.ReplaceNotes(fresh, "refilled"); err != nil {
			t.Fatal(err)
		}
	})
	t.Run("part fingerprint catches direct low level change", func(t *testing.T) {
		s, _ := openNotesFixture(t)
		target, err := s.FindNotes(slide)
		if err != nil {
			t.Fatal(err)
		}
		data, hash, _ := s.pkg.Part(part)
		data = bytes.Replace(data, []byte("Gothic"), []byte("Victorian"), 1)
		if err = s.pkg.Replace([]packaging.Replacement{{Part: part, ExpectedSHA256: hash, Data: data}}); err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceNotes(target, "stale"); err == nil {
			t.Fatal("changed part fingerprint accepted")
		}
	})
}
