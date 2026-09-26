package packaging

import (
	"bytes"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// This is one corpus batch: all checked-in Office fixtures, including duplicate
// content with different provenance paths. No graph/format-edit parity is implied.
func TestRetainedNoOpFixtureCorpus(t *testing.T) {
	root := testutil.FixturePath()
	count := 0
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".docx" && ext != ".xlsx" && ext != ".pptx" {
			return nil
		}
		count++
		t.Run(filepath.ToSlash(path), func(t *testing.T) {
			source, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			p, err := OpenPreserved(source, Limits{MaxSourceBytes: 64 << 20, MaxEntries: 4096, MaxPartBytes: 32 << 20, MaxTotalBytes: 128 << 20})
			if err != nil {
				t.Fatal(err)
			}
			var b bytes.Buffer
			if err = p.WriteTo(&b); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(source, b.Bytes()) {
				t.Fatal("no-op changed fixture bytes")
			}
		})
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	if count == 0 {
		t.Fatal("empty corpus")
	}
	t.Logf("%d Office fixtures tested", count)
}
