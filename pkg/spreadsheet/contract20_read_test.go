package spreadsheet

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

const contract20ReadCommit = "f4b5be7998ae6d009429f2080b40e3d4530c3452"
const contract20ReadManifest = "e8a6fc090257f64d2a1487edbfea0b151ff0aff66de35d4d0d098d3b35284aa7"
const contract20ReadRecipe = "9587da306418a4e67e7f23304beacdd5adec49719a7319342879bc315396c6fa"

// A native test input is packed from literal shared recipe members, never
// emitted by a production workbook writer. No generated archive hash is an
// acceptance gate: independent reopened values and physical input custody are.
func contract20ReadInput(t *testing.T, id string) []byte {
	t.Helper()
	root := os.Getenv("OOXML_FIXTURES_ROOT")
	pinPath := os.Getenv("OOXML_REFERENCE_PIN")
	if root == "" || pinPath == "" {
		t.Skip("explicit Contract20 root/pin required")
	}
	var pin struct {
		Schema   int    `json:"schema"`
		Commit   string `json:"commit"`
		Tag      string `json:"tag"`
		Manifest string `json:"manifest_sha256"`
	}
	b, err := os.ReadFile(pinPath)
	if err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(b, &pin); err != nil {
		t.Fatal(err)
	}
	if pin.Schema != 2 || pin.Commit != contract20ReadCommit || pin.Tag != "candidate-contract20" || pin.Manifest != contract20ReadManifest || root != "/workspace/projects/fixtures-ooxml" {
		t.Skip("Contract20 candidate not selected")
	}
	if err := testutil.VerifyReferenceCheckout(root, testutil.ReferenceIdentity{Schema: 2, Commit: pin.Commit, Tag: pin.Tag, Manifest: pin.Manifest}, true); err != nil {
		t.Fatal(err)
	}
	rel := filepath.Join("ledgers", "contract20-recipes.json")
	recipe, err := os.ReadFile(filepath.Join(root, rel))
	if err != nil {
		t.Fatal(err)
	}
	h := sha256.Sum256(recipe)
	if hex.EncodeToString(h[:]) != contract20ReadRecipe {
		t.Fatal("Contract20 recipe provenance drift")
	}
	var pack struct {
		XLSX []struct {
			ID      string            `json:"id"`
			Members map[string]string `json:"members"`
		} `json:"xlsx"`
		ExecutionCredit bool `json:"executionCredit"`
	}
	if err := json.Unmarshal(recipe, &pack); err != nil {
		t.Fatal(err)
	}
	if pack.ExecutionCredit {
		t.Fatal("recipe claims consumer execution credit")
	}
	var members map[string]string
	for _, row := range pack.XLSX {
		if row.ID == id {
			members = row.Members
			break
		}
	}
	if members == nil {
		t.Fatalf("missing literal recipe %s", id)
	}
	var out bytes.Buffer
	w := zip.NewWriter(&out)
	keys := make([]string, 0, len(members))
	for name := range members {
		keys = append(keys, name)
	}
	// Sorting makes fixture construction reproducible, not a generated-byte oracle.
	sort.Strings(keys)
	for _, name := range keys {
		if !filepath.IsLocal(name) || strings.Contains(name, "\\") {
			t.Fatal("unsafe literal member", name)
		}
		f, e := w.Create(name)
		if e != nil {
			t.Fatal(e)
		}
		if _, e = f.Write([]byte(members[name])); e != nil {
			t.Fatal(e)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func contract20InputMembers(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	z, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	out := map[string][]byte{}
	for _, file := range z.File {
		if _, exists := out[file.Name]; exists {
			t.Fatal("duplicate member", file.Name)
		}
		r, e := file.Open()
		if e != nil {
			t.Fatal(e)
		}
		content, e := io.ReadAll(r)
		closeErr := r.Close()
		if e != nil || closeErr != nil {
			t.Fatal(e, closeErr)
		}
		out[file.Name] = content
	}
	return out
}

func TestContract20XLSXReadProfiles(t *testing.T) {
	t.Run("relationship linked sheets and typed values", func(t *testing.T) {
		data := contract20ReadInput(t, "rel-values")
		original := append([]byte(nil), data...)
		parts := contract20InputMembers(t, data)
		wb, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			t.Fatal(err)
		}
		defer wb.Close()
		if wb.SheetCount() != 2 {
			t.Fatalf("sheet count = %d", wb.SheetCount())
		}
		for i, want := range []string{"Alpha", "Beta"} {
			sheet, e := wb.Sheet(i)
			if e != nil {
				t.Fatal(e)
			}
			if sheet.Name() != want {
				t.Fatalf("sheet %d = %q", i, sheet.Name())
			}
		}
		for _, row := range []struct {
			name, shared   string
			number         float64
			boolean        bool
			formula, cache string
		}{{"Alpha", "alpha shared", 7, false, "B1+1", "8"}, {"Beta", "beta shared", 42, true, "B1*2", "84"}} {
			sheet, e := wb.Sheet(row.name)
			if e != nil {
				t.Fatal(e)
			}
			if got := sheet.Cell("A1").String(); got != row.shared {
				t.Errorf("%s A1 = %q", row.name, got)
			}
			if got, e := sheet.Cell("B1").Float64(); e != nil || got != row.number {
				t.Errorf("%s B1 = %v / %v", row.name, got, e)
			}
			if got, e := sheet.Cell("C1").Bool(); e != nil || got != row.boolean {
				t.Errorf("%s C1 = %v / %v", row.name, got, e)
			}
			if got := sheet.Cell("D1").Formula(); got != row.formula {
				t.Errorf("%s D1 formula = %q", row.name, got)
			}
			if got := sheet.Cell("D1").String(); got != row.cache {
				t.Errorf("%s D1 cache = %q", row.name, got)
			}
		}
		if !bytes.Equal(data, original) || !reflect.DeepEqual(parts, contract20InputMembers(t, data)) {
			t.Error("input archive/member custody changed")
		}
		// The two workbook edges are explicit; index-based physical sheet order is not a substitute.
		rels := string(parts["xl/_rels/workbook.xml.rels"])
		workbook := string(parts["xl/workbook.xml"])
		for _, literal := range []string{`name="Alpha" sheetId="1" r:id="rId2"`, `name="Beta" sheetId="2" r:id="rId1"`} {
			if !strings.Contains(workbook, literal) {
				t.Errorf("missing workbook edge %s", literal)
			}
		}
		for _, literal := range []string{`Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"`, `Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"`} {
			if !strings.Contains(rels, literal) {
				t.Errorf("missing relationship %s", literal)
			}
		}
	})
	t.Run("visible rich runs exclude phonetic guide", func(t *testing.T) {
		data := contract20ReadInput(t, "phonetic-strings")
		original := append([]byte(nil), data...)
		parts := contract20InputMembers(t, data)
		wb, err := OpenReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			t.Fatal(err)
		}
		defer wb.Close()
		sheet, e := wb.Sheet(0)
		if e != nil {
			t.Fatal(e)
		}
		for _, item := range []struct{ cell, want string }{{"A1", "Alpha Beta"}, {"B1", "Gamma Delta"}} {
			c := sheet.Cell(item.cell)
			if c == nil {
				t.Fatalf("missing %s", item.cell)
			}
			if got := c.String(); got != item.want {
				t.Errorf("%s visible text = %q, want %q", item.cell, got, item.want)
			}
		}
		if !bytes.Equal(data, original) || !reflect.DeepEqual(parts, contract20InputMembers(t, data)) {
			t.Error("read changed input archive/member bytes")
		}
		if string(parts["custom/custody.bin"]) != "contract20-unchanged" {
			t.Error("custody payload drift")
		}
		for _, name := range []string{"xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"} {
			xml := string(parts[name])
			if !strings.Contains(xml, `xml:space="preserve"`) || !strings.Contains(xml, `<rPh sb="0" eb="5"><t>Guide</t></rPh>`) {
				t.Errorf("physical rich/phonetic structure missing in %s", name)
			}
		}
	})
}
