package acceptance

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Native alternates are controls, not additional Contract20 case credit.
func TestContract20TitleCollateralBatch(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	t.Run("alternate-cross-run-range", func(t *testing.T) {
		w := &contract20TitleWorld{temp: t.TempDir()}
		if err := w.input("fragmented-title"); err != nil {
			t.Fatal(err)
		}
		if err := w.find(); err != nil {
			t.Fatal(err)
		}
		if w.anchor.Text() != "Frankenstein" {
			t.Fatalf("original title %q", w.anchor.Text())
		}
		changed, err := w.session.ReplaceContractTitleSpan(w.anchor, 4, 7, "iend")
		if err != nil || !changed {
			t.Fatalf("alternate edit %v/%v", changed, err)
		}
		path := filepath.Join(w.temp, "alternate.pptx")
		if _, err := w.session.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		archive, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		members, err := blankMemberPayloads(archive)
		if err != nil {
			t.Fatal(err)
		}
		want := []presentation.ContractTitleFragment{{Text: "Fran", Attributes: map[string]string{"b": "1"}}, {Text: "iend", Attributes: map[string]string{"i": "1"}}, {Text: "stein", Attributes: map[string]string{"u": "sng"}}}
		runs, err := oracleContract20TitleRuns(members["ppt/slides/slide1.xml"])
		if err != nil || !reflect.DeepEqual(runs, want) {
			t.Fatalf("alternate runs %+v: %v", runs, err)
		}
		const old = `<a:p><a:r><a:rPr b="1"/><a:t>Fran</a:t></a:r><a:r><a:rPr i="1"/><a:t>ken</a:t></a:r><a:r><a:rPr u="sng"/><a:t>stein</a:t></a:r></a:p>`
		const expected = `<a:p><a:r><a:rPr b="1"/><a:t>Fran</a:t></a:r><a:r><a:rPr i="1"/><a:t>iend</a:t></a:r><a:r><a:rPr u="sng"/><a:t>stein</a:t></a:r></a:p>`
		before := w.before["ppt/slides/slide1.xml"]
		after := members["ppt/slides/slide1.xml"]
		if strings.Count(string(before), old) != 1 || !bytes.Equal(after, []byte(strings.Replace(string(before), old, expected, 1))) {
			t.Fatal("alternate lexical mask mismatch")
		}
		if len(members) != len(w.before) {
			t.Fatal("member inventory changed")
		}
		for part, original := range w.before {
			if part != "ppt/slides/slide1.xml" && !bytes.Equal(members[part], original) {
				t.Fatalf("unrelated member %s", part)
			}
		}
		graph, err := contract20TableGraph(archive)
		if err != nil || !reflect.DeepEqual(graph, w.graph) {
			t.Fatalf("OPC graph changed: %v", err)
		}
		if err := w.unchanged(); err != nil {
			t.Fatal(err)
		}
	})
	for _, tc := range []struct{ name, separator string }{
		{"inter-run-whitespace", " \n\t "},
		{"inter-run-comment", `<!--keep sibling-->`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			w := &contract20TitleWorld{temp: t.TempDir()}
			if err := w.input("fragmented-title"); err != nil {
				t.Fatal(err)
			}
			parts := make(map[string][]byte, len(w.before))
			for part, data := range w.before {
				parts[part] = bytes.Clone(data)
			}
			const part = "ppt/slides/slide1.xml"
			const boundary = `<a:t>Fran</a:t></a:r><a:r><a:rPr i="1"/>`
			if strings.Count(string(parts[part]), boundary) != 1 {
				t.Fatal("original title run boundary not unique")
			}
			parts[part] = []byte(strings.Replace(string(parts[part]), boundary, `<a:t>Fran</a:t></a:r>`+tc.separator+`<a:r><a:rPr i="1"/>`, 1))
			archive, err := contract20PackMembers(parts)
			if err != nil {
				t.Fatal(err)
			}
			original := bytes.Clone(archive)
			graph, err := contract20TableGraph(archive)
			if err != nil {
				t.Fatal(err)
			}
			s, err := presentation.OpenEditing(archive, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			anchor, err := s.FindContractTitle(part)
			if err != nil || anchor.Text() != "Frankenstein" {
				t.Fatalf("source title %q: %v", anchor.Text(), err)
			}
			changed, err := s.ReplaceContractTitleSpan(anchor, 2, 7, "iend")
			var refusal *packaging.Refusal
			if changed || !errors.As(err, &refusal) || refusal.Kind != "PPTX_UNSUPPORTED_TEXT_TOPOLOGY" {
				t.Fatalf("inter-run lexical sibling was lost or accepted: %v/%v", changed, err)
			}
			path := filepath.Join(w.temp, "refused.pptx")
			if _, err := s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			after, err := blankMemberPayloads(saved)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(archive, original) || !reflect.DeepEqual(parts, after) {
				t.Fatal("refusal changed caller or complete saved member state")
			}
			afterGraph, err := contract20TableGraph(saved)
			if err != nil || !reflect.DeepEqual(graph, afterGraph) {
				t.Fatalf("refusal changed saved OPC graph: %v", err)
			}
			fresh, err := s.FindContractTitle(part)
			if err != nil || fresh.Text() != anchor.Text() {
				t.Fatalf("held title invalidated on refusal %q: %v", fresh.Text(), err)
			}
			if err := w.unchanged(); err != nil {
				t.Fatal(err)
			}
		})
	}
	t.Run("split-utf16-surrogate-refusal", func(t *testing.T) {
		w := &contract20TitleWorld{temp: t.TempDir()}
		if err := w.input("plain-title"); err != nil {
			t.Fatal(err)
		}
		parts := make(map[string][]byte, len(w.before))
		for name, data := range w.before {
			parts[name] = bytes.Clone(data)
		}
		const part = "ppt/slides/slide1.xml"
		if strings.Count(string(parts[part]), `<a:t>Frankenstein</a:t>`) != 1 {
			t.Fatal("source title anchor not unique")
		}
		parts[part] = []byte(strings.Replace(string(parts[part]), `<a:t>Frankenstein</a:t>`, `<a:t>F😀ankenstein</a:t>`, 1))
		archive, err := contract20PackMembers(parts)
		if err != nil {
			t.Fatal(err)
		}
		original := bytes.Clone(archive)
		graph, err := contract20TableGraph(archive)
		if err != nil {
			t.Fatal(err)
		}
		s, err := presentation.OpenEditing(archive, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		anchor, err := s.FindContractTitle(part)
		if err != nil || anchor.Text() != "F😀ankenstein" {
			t.Fatalf("UTF-16 source title %q: %v", anchor.Text(), err)
		}
		changed, err := s.ReplaceContractTitleSpan(anchor, 2, 3, "X")
		var refusal *packaging.Refusal
		if changed || !errors.As(err, &refusal) || refusal.Kind != "PPTX_UNSUPPORTED_TEXT_TOPOLOGY" {
			t.Fatalf("split surrogate accepted %v/%v", changed, err)
		}
		path := filepath.Join(w.temp, "unchanged.pptx")
		if _, err := s.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		saved, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		after, err := blankMemberPayloads(saved)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(archive, original) || !reflect.DeepEqual(parts, after) {
			t.Fatal("split refusal changed caller or saved member state")
		}
		afterGraph, err := contract20TableGraph(saved)
		if err != nil || !reflect.DeepEqual(graph, afterGraph) {
			t.Fatalf("split refusal changed graph: %v", err)
		}
		fresh, err := s.FindContractTitle(part)
		if err != nil || fresh.Text() != anchor.Text() {
			t.Fatalf("refusal invalidated title %q: %v", fresh.Text(), err)
		}
		if err := w.unchanged(); err != nil {
			t.Fatal(err)
		}
	})
}

func TestContract20TitleIssuanceCollateral(t *testing.T) {
	if !contract20Candidate() {
		t.Skip("exact Contract20 candidate not selected")
	}
	loadReferencePin(t)
	w := &contract20TitleWorld{temp: t.TempDir()}
	if err := w.input("plain-title"); err != nil {
		t.Fatal(err)
	}
	if err := w.find(); err != nil {
		t.Fatal(err)
	}
	foreign := &contract20TitleWorld{temp: t.TempDir()}
	if err := foreign.input("plain-title"); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		label  string
		s      *presentation.EditSession
		target *presentation.ContractTitleAnchor
	}{{"foreign", foreign.session, w.anchor}, {"copied", w.session, func() *presentation.ContractTitleAnchor { copy := *w.anchor; return &copy }()}, {"nil", w.session, nil}} {
		changed, err := tc.s.ReplaceContractTitleSpan(tc.target, 0, 12, "Creature")
		var refusal *packaging.Refusal
		if changed || !errors.As(err, &refusal) || refusal.Kind != "PPTX_STALE_ANCHOR" {
			t.Fatalf("%s accepted %v/%v", tc.label, changed, err)
		}
	}
	if err := w.unchanged(); err != nil {
		t.Fatal(err)
	}
	changed, err := w.session.ReplaceContractTitleSpan(w.anchor, 0, 12, "Creature")
	if !changed || err != nil {
		t.Fatal(fmt.Errorf("original anchor lost after refusal %v/%v", changed, err))
	}
	path := filepath.Join(w.temp, "issued.pptx")
	if _, err := w.session.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	archive, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	parts, err := blankMemberPayloads(archive)
	if err != nil {
		t.Fatal(err)
	}
	title, err := oracleContract20Title(parts["ppt/slides/slide1.xml"])
	if err != nil || title != "Creature" {
		t.Fatalf("after issuance title %q: %v", title, err)
	}
}
