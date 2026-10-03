package presentation

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const contract20TableCommit = "f4b5be7998ae6d009429f2080b40e3d4530c3452"
const contract20TableManifest = "e8a6fc090257f64d2a1487edbfea0b151ff0aff66de35d4d0d098d3b35284aa7"
const contract20TableRecipes = "9587da306418a4e67e7f23304beacdd5adec49719a7319342879bc315396c6fa"

func contract20PPTXTableInput(t *testing.T, id string) []byte {
	t.Helper()
	root, pinPath := os.Getenv("OOXML_FIXTURES_ROOT"), os.Getenv("OOXML_REFERENCE_PIN")
	if root == "" || pinPath == "" {
		t.Skip("explicit Contract20 candidate required")
	}
	var pin struct {
		Schema   int    `json:"schema"`
		Commit   string `json:"commit"`
		Tag      string `json:"tag"`
		Manifest string `json:"manifest_sha256"`
	}
	pinBytes, err := os.ReadFile(pinPath)
	if err != nil {
		t.Fatal(err)
	}
	if err = json.Unmarshal(pinBytes, &pin); err != nil {
		t.Fatal(err)
	}
	if pin.Schema != 2 || pin.Commit != contract20TableCommit || pin.Tag != "candidate-contract20" || pin.Manifest != contract20TableManifest || root != "/workspace/projects/fixtures-ooxml" {
		t.Skip("different candidate")
	}
	if err = testutil.VerifyReferenceCheckout(root, testutil.ReferenceIdentity{Schema: 2, Commit: pin.Commit, Tag: pin.Tag, Manifest: pin.Manifest}, true); err != nil {
		t.Fatal(err)
	}
	recipeBytes, err := os.ReadFile(testutil.ReferencePath("ledgers", "contract20-recipes.json"))
	if err != nil {
		t.Fatal(err)
	}
	hash := sha256.Sum256(recipeBytes)
	if hex.EncodeToString(hash[:]) != contract20TableRecipes {
		t.Fatal("recipe seal differs")
	}
	var recipes struct {
		PPTX []struct {
			ID            string `json:"id"`
			BaseFixtureID string `json:"baseFixtureId"`
			Operations    []struct {
				Kind   string `json:"kind"`
				Part   string `json:"part"`
				Before string `json:"before"`
				After  string `json:"after"`
			} `json:"operations"`
		} `json:"pptx"`
	}
	if err = json.Unmarshal(recipeBytes, &recipes); err != nil {
		t.Fatal(err)
	}
	for _, r := range recipes.PPTX {
		if r.ID != id {
			continue
		}
		sourcePath, err := testutil.LookupFixture(r.BaseFixtureID)
		if err != nil {
			t.Fatal(err)
		}
		source, err := os.ReadFile(sourcePath)
		if err != nil {
			t.Fatal(err)
		}
		zr, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
		if err != nil {
			t.Fatal(err)
		}
		members := map[string][]byte{}
		for _, f := range zr.File {
			if _, exists := members[f.Name]; exists {
				t.Fatal("duplicate source member")
			}
			stream, e := f.Open()
			if e != nil {
				t.Fatal(e)
			}
			b, e := io.ReadAll(stream)
			closeErr := stream.Close()
			if e != nil || closeErr != nil {
				t.Fatal(e, closeErr)
			}
			members[f.Name] = b
		}
		for _, op := range r.Operations {
			if op.Kind != "replace-literal-once" {
				t.Fatalf("unsupported recipe operation %s", op.Kind)
			}
			original, ok := members[op.Part]
			if !ok || strings.Count(string(original), op.Before) != 1 {
				t.Fatalf("unsealed literal replacement %s", op.Part)
			}
			members[op.Part] = []byte(strings.Replace(string(original), op.Before, op.After, 1))
		}
		var out bytes.Buffer
		zw := zip.NewWriter(&out)
		names := make([]string, 0, len(members))
		for n := range members {
			names = append(names, n)
		}
		sort.Strings(names)
		for _, n := range names {
			w, e := zw.Create(n)
			if e != nil {
				t.Fatal(e)
			}
			if _, e = w.Write(members[n]); e != nil {
				t.Fatal(e)
			}
		}
		if err = zw.Close(); err != nil {
			t.Fatal(err)
		}
		return out.Bytes()
	}
	t.Fatalf("missing PPTX recipe %s", id)
	return nil
}
func contract20PPTXMembers(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	z, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	out := map[string][]byte{}
	for _, f := range z.File {
		if _, ok := out[f.Name]; ok {
			t.Fatal("duplicate member")
		}
		r, e := f.Open()
		if e != nil {
			t.Fatal(e)
		}
		b, e := io.ReadAll(r)
		closeErr := r.Close()
		if e != nil || closeErr != nil {
			t.Fatal(e, closeErr)
		}
		out[f.Name] = b
	}
	return out
}
func TestContract20TableCellEditBatch(t *testing.T) {
	for _, v := range []struct{ recipe, want string }{{"merged-table", "PPTX_TABLE_MERGE_UNSUPPORTED"}, {"malformed-table", "PPTX_TABLE_STRUCTURE_UNSUPPORTED"}, {"styled-table", ""}} {
		t.Run(v.recipe, func(t *testing.T) {
			source := contract20PPTXTableInput(t, v.recipe)
			original := append([]byte(nil), source...)
			members := contract20PPTXMembers(t, source)
			session, err := OpenEditing(source, packaging.Limits{})
			if err != nil {
				t.Fatalf("valid package must OPEN: %v", err)
			}
			cell, err := session.FindContractTableCell("ppt/slides/slide1.xml", 4, 0, 0)
			if err != nil {
				t.Fatalf("valid frame/cell must SELECT: %v", err)
			}
			err = session.SetContractTableCellText(cell, "updated value")
			if v.want != "" {
				var refusal *packaging.Refusal
				if !errors.As(err, &refusal) || refusal.Kind != v.want {
					t.Errorf("edit refusal = %v, want %s", err, v.want)
				}
				dest := filepath.Join(t.TempDir(), "refused.pptx")
				if _, err = session.SaveAs(dest); err != nil {
					t.Fatal(err)
				}
				saved, err := os.ReadFile(dest)
				if err != nil {
					t.Fatal(err)
				}
				if !reflect.DeepEqual(members, contract20PPTXMembers(t, saved)) {
					t.Error("refused cell edit changed source members")
				}
			} else {
				if err != nil {
					t.Fatal(err)
				}
				dest := filepath.Join(t.TempDir(), "edited.pptx")
				if _, err = session.SaveAs(dest); err != nil {
					t.Fatal(err)
				}
				saved, err := os.ReadFile(dest)
				if err != nil {
					t.Fatal(err)
				}
				after := contract20PPTXMembers(t, saved)
				for part, was := range members {
					if part != "ppt/slides/slide1.xml" && !bytes.Equal(was, after[part]) {
						t.Errorf("unrelated member changed %s", part)
					}
				}
				oldSlide := string(members["ppt/slides/slide1.xml"])
				newSlide := string(after["ppt/slides/slide1.xml"])
				if strings.Count(oldSlide, "Galvanic battery") != 1 || newSlide != strings.Replace(oldSlide, "Galvanic battery", "updated value", 1) {
					t.Error("styled edit changed bytes outside selected text leaf")
				}
				reopened, err := OpenEditing(saved, packaging.Limits{})
				if err != nil {
					t.Fatalf("saved package failed production reopen: %v", err)
				}
				if _, err := reopened.FindContractTableCell("ppt/slides/slide1.xml", 4, 0, 0); err != nil {
					t.Fatalf("reopened cell unavailable: %v", err)
				}
				stale := session.SetContractTableCellText(cell, "again")
				var refusal *packaging.Refusal
				if !errors.As(stale, &refusal) || refusal.Kind != "PPTX_STALE_TABLE_HANDLE" {
					t.Errorf("consumed cell = %v", stale)
				}
			}
			if !bytes.Equal(source, original) || !reflect.DeepEqual(members, contract20PPTXMembers(t, source)) {
				t.Error("caller/source archive modified")
			}
		})
	}
}
