package spreadsheet

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// Transform only private literal-recipe archives. Golden expectations remain
// independent XML/graph/member observations rather than generated ZIP hashes.
func contract20AlterMember(t *testing.T, input []byte, name string, alter func(string) string) []byte {
	t.Helper()
	r, err := zip.NewReader(bytes.NewReader(input), int64(len(input)))
	if err != nil {
		t.Fatal(err)
	}
	members := map[string][]byte{}
	for _, f := range r.File {
		body, e := f.Open()
		if e != nil {
			t.Fatal(e)
		}
		data, e := io.ReadAll(body)
		closeErr := body.Close()
		if e != nil || closeErr != nil {
			t.Fatal(e, closeErr)
		}
		members[f.Name] = data
	}
	old, ok := members[name]
	if !ok {
		t.Fatalf("missing member %s", name)
	}
	changed := alter(string(old))
	if changed == string(old) {
		t.Fatal("negative control made no change")
	}
	members[name] = []byte(changed)
	var out bytes.Buffer
	w := zip.NewWriter(&out)
	names := make([]string, 0, len(members))
	for part := range members {
		names = append(names, part)
	}
	sort.Strings(names)
	for _, part := range names {
		writer, e := w.Create(part)
		if e != nil {
			t.Fatal(e)
		}
		if _, e = writer.Write(members[part]); e != nil {
			t.Fatal(e)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func TestContract20UnselectedTopologySafetyBatch(t *testing.T) {
	for _, variant := range []struct{ name, relationship string }{
		{"unreferenced worksheet edge at workbook", `<Relationship Id="extraSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`},
		{"styles edge aimed at workbook", `<Relationship Id="extraStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="workbook.xml"/>`},
		{"styles edge aimed at custom payload", `<Relationship Id="extraStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../custom/custody.bin"/>`},
		{"shared strings edge aimed at opaque cache", `<Relationship Id="extraShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="charts/cache-boundary.xml"/>`},
		{"theme edge aimed at custom payload", `<Relationship Id="extraTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../custom/custody.bin"/>`},
	} {
		t.Run(variant.name, func(t *testing.T) {
			input := contract20ReadInput(t, "opaque-caches")
			modified := contract20AlterMember(t, input, "xl/_rels/workbook.xml.rels", func(old string) string {
				return strings.Replace(old, "</Relationships>", variant.relationship+"</Relationships>", 1)
			})
			before := contract20InputMembers(t, modified)
			s, err := OpenEditing(modified, packaging.Limits{})
			if err != nil {
				t.Fatal(err)
			}
			target, err := s.FindNumber("Model", "A1")
			if err != nil {
				t.Fatal(err)
			}
			effect, err := s.SetNumberWithSealedOpaqueCaches(target, 10)
			var refusal *packaging.Refusal
			if !errors.As(err, &refusal) || refusal.Kind != "xlsx-opaque-cache-scope-unsupported" || effect.ValueChanged {
				t.Errorf("owner/target refusal %#v, effect %#v", err, effect)
			}
			if value, err := s.CurrentNumber(target); err != nil || value != 1 {
				t.Errorf("held target changed %v/%v", value, err)
			}
			path := filepath.Join(t.TempDir(), "refused.xlsx")
			if _, err := s.SaveAs(path); err != nil {
				t.Fatal(err)
			}
			data, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(before, contract20InputMembers(t, data)) {
				t.Error("owner/target refusal changed members")
			}
		})
	}
	t.Run("extra root worksheet edge refuses inert-cache opt-in", func(t *testing.T) {
		input := contract20ReadInput(t, "opaque-caches")
		variant := contract20AlterMember(t, input, "_rels/.rels", func(old string) string {
			return strings.Replace(old, "</Relationships>", `<Relationship Id="rootSemantic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="xl/worksheets/sheet1.xml"/></Relationships>`, 1)
		})
		before := contract20InputMembers(t, variant)
		s, err := OpenEditing(variant, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		target, err := s.FindNumber("Model", "A1")
		if err != nil {
			t.Fatal(err)
		}
		effect, err := s.SetNumberWithSealedOpaqueCaches(target, 10)
		var refusal *packaging.Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "xlsx-opaque-cache-scope-unsupported" || effect.ValueChanged {
			t.Errorf("root edge refusal %#v, effect %#v", err, effect)
		}
		if got, err := s.CurrentNumber(target); err != nil || got != 1 {
			t.Errorf("held input consumed by root refusal: %v/%v", got, err)
		}
		path := filepath.Join(t.TempDir(), "after-refusal.xlsx")
		if _, err := s.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		data, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		if !reflect.DeepEqual(before, contract20InputMembers(t, data)) {
			t.Error("refused root variant changed saved members")
		}
	})
	t.Run("shared formula anywhere blocks changed input before sweep", func(t *testing.T) {
		input := contract20ReadInput(t, "cross-caches")
		variant := contract20AlterMember(t, input, "xl/worksheets/sheet2.xml", func(old string) string {
			return strings.Replace(old, "<f>40+2</f>", `<f t="shared" si="0" ref="B1:B2">40+2</f>`, 1)
		})
		before := contract20InputMembers(t, variant)
		s, err := OpenEditing(variant, packaging.Limits{})
		if err != nil {
			t.Fatal(err)
		}
		target, err := s.FindNumber("Model", "A1")
		if err != nil {
			t.Fatal(err)
		}
		if effect, err := s.SetNumberWithAllCachesInvalidated(target, 1); err != nil || effect.ValueChanged {
			t.Errorf("same numeric no-op = %#v/%v", effect, err)
		}
		effect, err := s.SetNumberWithAllCachesInvalidated(target, 10)
		var refusal *packaging.Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "xlsx-cache-topology-unsupported" || effect.ValueChanged {
			t.Errorf("shared topology refusal %#v, effect %#v", err, effect)
		}
		if got, err := s.CurrentNumber(target); err != nil || got != 1 {
			t.Errorf("held input consumed by shared refusal: %v/%v", got, err)
		}
		path := filepath.Join(t.TempDir(), "after-refusal.xlsx")
		if _, err := s.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		data, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		if !reflect.DeepEqual(before, contract20InputMembers(t, data)) {
			t.Error("refused shared variant changed saved members")
		}
	})
}
