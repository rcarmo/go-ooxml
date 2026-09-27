package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func TestPackageSerializationContracts(t *testing.T) {
	t.Run("closed stream", func(t *testing.T) {
		p := New()
		_ = p.Close()
		var b bytes.Buffer
		if err := p.WriteTo(&b); !errors.Is(err, utils.ErrDocumentClosed) {
			t.Fatalf("got %v", err)
		}
		if b.Len() != 0 {
			t.Fatal("closed package wrote bytes")
		}
	})
	t.Run("ordered members", func(t *testing.T) {
		p := New()
		for _, name := range []string{"z.xml", "a.xml", "k.xml", "b.xml"} {
			_, _ = p.AddPart(name, ContentTypeXML, []byte(`<data/>`))
			p.AddRelationship(name, "https://example.org", RelTypeHyperlink)
		}
		var reference []byte
		for i := 0; i < 5; i++ {
			var b bytes.Buffer
			if err := p.WriteTo(&b); err != nil {
				t.Fatal(err)
			}
			if i > 0 && !bytes.Equal(reference, b.Bytes()) {
				t.Fatal("same state serialized differently")
			}
			reference = bytes.Clone(b.Bytes())
			r, err := zip.NewReader(bytes.NewReader(b.Bytes()), int64(b.Len()))
			if err != nil {
				t.Fatal(err)
			}
			var names []string
			for _, f := range r.File {
				names = append(names, f.Name)
			}
			// Registries precede ordinary parts; each group is lexical.
			rels := names[1:5]
			if !sort.StringsAreSorted(rels) || !sort.StringsAreSorted(names[5:]) {
				t.Fatalf("unstable member order %v", names)
			}
		}
	})
	t.Run("atomic success and rename failure", func(t *testing.T) {
		dir := t.TempDir()
		dest := filepath.Join(dir, "doc.docx")
		p := New()
		_, _ = p.AddPart("a.xml", ContentTypeXML, []byte(`<data/>`))
		if err := os.WriteFile(dest, []byte("old"), 0640); err != nil {
			t.Fatal(err)
		}
		if err := p.SaveAs(dest); err != nil {
			t.Fatal(err)
		}
		if p.Path() != dest || p.IsModified() {
			t.Fatal("success state not committed")
		}
		st, err := os.Stat(dest)
		if err != nil {
			t.Fatal(err)
		}
		if st.Mode().Perm() != 0640 {
			t.Fatalf("permissions %v", st.Mode())
		}
		if q, err := Open(dest); err != nil {
			t.Fatal(err)
		} else {
			_ = q.Close()
		}
		blocked := filepath.Join(dir, "directory")
		if err = os.Mkdir(blocked, 0700); err != nil {
			t.Fatal(err)
		}
		if err = p.SaveAs(blocked); err == nil {
			t.Fatal("directory destination accepted")
		}
		if p.Path() != dest {
			t.Fatal("failed save changed path")
		}
		files, _ := os.ReadDir(dir)
		var names []string
		for _, f := range files {
			names = append(names, f.Name())
		}
		if !reflect.DeepEqual(names, []string{"directory", "doc.docx"}) {
			t.Fatalf("temporary files: %v", names)
		}
	})
}
