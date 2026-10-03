package acceptance

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
	"github.com/rcarmo/go-ooxml/pkg/presentation"
)

// Native layout admission controls, not additional selected case credit.
func TestContract20CreationLayoutAdmissionControls(t *testing.T) {
	contract20CreationSelected(t, "@id-pptx-create-refusals")
	for _, tc := range []struct {
		name, before, after string
		accept              bool
		subtitleIdx         string
	}{
		{"subtitle-index-seven", `<p:ph type="subTitle" idx="1"/>`, `<p:ph type="subTitle" idx="7"/>`, true, "7"},
		{"title-type-direct", `<p:ph type="ctrTitle"/>`, `<p:ph type="title"/>`, true, "1"},
		{"subtitle-leading-zero-index", `<p:ph type="subTitle" idx="1"/>`, `<p:ph type="subTitle" idx="07"/>`, false, ""},
		{"subtitle-missing-index", `<p:ph type="subTitle" idx="1"/>`, `<p:ph type="subTitle"/>`, false, ""},
		{"duplicate-shape-identity", `<p:cNvPr id="3" name="Subtitle 2"/>`, `<p:cNvPr id="2" name="Subtitle 2"/>`, false, ""},
		{"subtitle-wrong-owner", `<p:ph type="subTitle" idx="1"/>`, ``, false, ""},
		{"title-wrong-owner", `<p:ph type="ctrTitle"/>`, ``, false, ""},
	} {
		t.Run(tc.name, func(t *testing.T) {
			w := &contract20TitleWorld{temp: t.TempDir()}
			if err := w.input("title-subtitle-source"); err != nil {
				t.Fatal(err)
			}
			original := bytes.Clone(w.source)
			parts := map[string][]byte{}
			for name, data := range w.before {
				parts[name] = bytes.Clone(data)
			}
			const layout = "ppt/slideLayouts/slideLayout1.xml"
			old := parts[layout]
			if bytes.Count(old, []byte(tc.before)) != 1 {
				t.Fatalf("nonunique layout operand %q", tc.before)
			}
			updated := bytes.Replace(old, []byte(tc.before), []byte(tc.after), 1)
			if strings.HasSuffix(tc.name, "wrong-owner") {
				moved := `<p:extLst><p:ext uri="urn:contract20:wrong-owner"><p:nvPr>` + tc.before + `</p:nvPr></p:ext></p:extLst>`
				if bytes.Count(updated, []byte(`</p:sldLayout>`)) != 1 {
					t.Fatal("layout close not unique")
				}
				updated = bytes.Replace(updated, []byte(`</p:sldLayout>`), []byte(moved+`</p:sldLayout>`), 1)
			}
			parts[layout] = updated
			source, err := contract20PackMembers(parts)
			if err != nil {
				t.Fatal(err)
			}
			baseline, e := contract20TableGraph(source)
			if e != nil {
				t.Fatal(e)
			}
			session, e := presentation.OpenEditing(source, packaging.Limits{})
			if e != nil {
				t.Fatal(e)
			}
			order, e := creationSlideOrder(parts, baseline)
			if e != nil || len(order) != 1 {
				t.Fatalf("source order %q: %v", order, e)
			}
			held, e := session.FindContractTitle(order[0])
			if e != nil || held.Text() != "Original title" {
				t.Fatalf("original title anchor: %v", e)
			}
			existing := filepath.Join(t.TempDir(), "existing.pptx")
			sentinel := []byte("destination sentinel")
			if e = os.WriteFile(existing, sentinel, 0600); e != nil {
				t.Fatal(e)
			}
			e = session.AppendContractTitleSlide("Next title", "Next subtitle")
			if !tc.accept {
				var refusal *packaging.Refusal
				if !errors.As(e, &refusal) || refusal.Kind != "PPTX_LAYOUT_UNSAFE" {
					t.Fatalf("wrong owner admitted or wrong refusal: %v", e)
				}
				prior, e := os.ReadFile(existing)
				if e != nil || !bytes.Equal(prior, sentinel) {
					t.Fatalf("refusal changed destination: %v", e)
				}
				absent := filepath.Join(t.TempDir(), "absent.pptx")
				if _, e = os.Stat(absent); !os.IsNotExist(e) {
					t.Fatalf("absent path exists: %v", e)
				}
				restored := filepath.Join(t.TempDir(), "post-refusal.pptx")
				if _, e = session.SaveAs(restored); e != nil {
					t.Fatal(e)
				}
				saved, e := os.ReadFile(restored)
				if e != nil {
					t.Fatal(e)
				}
				got, e := blankMemberPayloads(saved)
				if e != nil {
					t.Fatal(e)
				}
				graph, e := contract20TableGraph(saved)
				if e != nil {
					t.Fatal(e)
				}
				if !reflect.DeepEqual(got, parts) || !reflect.DeepEqual(graph, baseline) || !bytes.Equal(w.source, original) {
					t.Fatal("refusal changed source/session members or relationships")
				}
				changed, e := session.ReplaceContractTitleSpan(held, 0, len("Original title"), held.Text())
				if e != nil || changed {
					t.Fatalf("held title no-op unusable after refusal: %v changed=%v", e, changed)
				}
				return
			}
			if e != nil {
				t.Fatal(e)
			}
			dest := filepath.Join(t.TempDir(), "accepted.pptx")
			if _, e = session.SaveAs(dest); e != nil {
				t.Fatal(e)
			}
			saved, e := os.ReadFile(dest)
			if e != nil {
				t.Fatal(e)
			}
			got, e := blankMemberPayloads(saved)
			if e != nil {
				t.Fatal(e)
			}
			graph, e := contract20TableGraph(saved)
			if e != nil {
				t.Fatal(e)
			}
			ordered, e := creationSlideOrder(got, graph)
			if e != nil || len(ordered) != 2 {
				t.Fatalf("saved order: %q %v", ordered, e)
			}
			if e = creationGraphDelta(baseline, graph, order[0], ordered[1]); e != nil {
				t.Fatal(e)
			}
			if !bytes.Equal(got[layout], updated) || !bytes.Equal(got[order[0]], parts[order[0]]) || !bytes.Equal(got[packaging.RelationshipsPathForPart(order[0])], parts[packaging.RelationshipsPathForPart(order[0])]) {
				t.Fatal("positive append changed old support/slide")
			}
			if strings.Count(string(got[ordered[1]]), `<p:ph type="subTitle" idx="`+tc.subtitleIdx+`"/>`) != 1 {
				t.Fatalf("new subtitle idx differs from admitted layout: %s", got[ordered[1]])
			}
			if tc.name == "title-type-direct" && strings.Count(string(got[ordered[1]]), `<p:ph type="title"/>`) != 1 {
				t.Fatal("new title placeholder type differs from admitted layout")
			}
			if e = creationPlaceholderOperands(got[ordered[1]], "Next title", "Next subtitle"); e != nil {
				t.Fatal(e)
			}
			changed, e := session.ReplaceContractTitleSpan(held, 0, len("Original title"), held.Text())
			if e != nil || changed {
				t.Fatalf("held title no-op unusable after successful append: %v changed=%v", e, changed)
			}
			if !bytes.Equal(w.source, original) {
				t.Fatal("caller archive mutated")
			}
		})
	}
}
