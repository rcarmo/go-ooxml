package packaging

import (
	"bytes"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// One logical corpus batch; duplicate labels resolve to one canonical file.
// These checks establish no-op custody, not graph/format-edit coverage.
func TestRetainedNoOpFixtureCorpus(t *testing.T) {
	count := 0
	for _, label := range testutil.FixtureLabels("") {
		ext := strings.ToLower(filepath.Ext(label))
		if ext != ".docx" && ext != ".xlsx" && ext != ".pptx" {
			continue
		}
		count++
		t.Run(label, func(t *testing.T) {
			path := testutil.FixturePath(label)
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
	}
	if count == 0 {
		t.Fatal("empty corpus")
	}
	t.Logf("%d Office fixtures tested", count)
}
