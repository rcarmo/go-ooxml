package presentation

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestRealFixtureRelationshipNamespaceEditBatch(t *testing.T) {
	fixture, err := testutil.LookupFixture("fixture-2aec94471f93c300d56ca4789106a974411085d1588f3424155362a06dd043f3")
	if err != nil {
		t.Fatal(err)
	}
	original, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name         string
		alias, wrong bool
	}{{name: "ordinary"}, {name: "alias", alias: true}, {name: "wrong URI", wrong: true}} {
		t.Run(tc.name, func(t *testing.T) {
			source := bytes.Clone(original)
			if tc.alias || tc.wrong {
				p, err := packaging.OpenPreserved(source, packaging.Limits{})
				if err != nil {
					t.Fatal(err)
				}
				main, hash, err := p.Part("ppt/presentation.xml")
				if err != nil {
					t.Fatal(err)
				}
				if tc.alias {
					main = bytes.Replace(main, []byte(`xmlns:r="`+packaging.NSDocumentRelationships+`"`), []byte(`xmlns:link="`+packaging.NSDocumentRelationships+`"`), 1)
					main = bytes.ReplaceAll(main, []byte(`r:id=`), []byte(`link:id=`))
				} else {
					main = bytes.Replace(main, []byte(`xmlns:r="`+packaging.NSDocumentRelationships+`"`), []byte(`xmlns:r="`+packaging.NSRelationships+`"`), 1)
				}
				if err = p.Replace([]packaging.Replacement{{Part: "ppt/presentation.xml", ExpectedSHA256: hash, Data: main}}); err != nil {
					t.Fatal(err)
				}
				path := filepath.Join(t.TempDir(), "input.pptx")
				if _, err = p.SaveAs(path); err != nil {
					t.Fatal(err)
				}
				source, err = os.ReadFile(path)
				if err != nil {
					t.Fatal(err)
				}
			}
			retained := bytes.Clone(source)
			s, err := OpenEditing(source, packaging.Limits{})
			if tc.wrong {
				var refused *packaging.Refusal
				if s != nil || !errors.As(err, &refused) || refused.Kind != "PPTX_PRESENTATION_INVALID" {
					t.Fatalf("wrong URI accepted or wrong reason: session=%v error=%v", s, err)
				}
				if !bytes.Equal(source, retained) {
					t.Fatal("caller bytes changed")
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			target, err := s.FindText("ppt/slides/slide1.xml", 2, "Original title")
			if err != nil {
				t.Fatal(err)
			}
			if err = s.Replace(target, "Edited title"); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(source, retained) {
				t.Fatal("caller bytes changed")
			}
			output := filepath.Join(t.TempDir(), "output.pptx")
			if _, err = s.SaveAs(output); err != nil {
				t.Fatal(err)
			}
			saved, err := os.ReadFile(output)
			if err != nil {
				t.Fatal(err)
			}
			again, err := OpenEditing(saved, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			if _, err = again.FindText("ppt/slides/slide1.xml", 2, "Edited title"); err != nil {
				t.Fatal(err)
			}
			graph, err := again.pkg.Graph()
			if err != nil {
				t.Fatal(err)
			}
			for _, e := range graph.Edges {
				if !e.External && e.ResolvedPart == "" {
					t.Fatalf("unresolved edge %+v", e)
				}
			}
			if tc.alias {
				main, _, err := again.pkg.Part("ppt/presentation.xml")
				if err != nil {
					t.Fatal(err)
				}
				if !strings.Contains(string(main), `link:id=`) || strings.Contains(string(main), `r:id=`) {
					t.Fatal("alias spelling lost")
				}
			}
		})
	}
}
