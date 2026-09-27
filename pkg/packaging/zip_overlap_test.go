package packaging

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func TestIndependentPhysicalOverlap(t *testing.T) {
	source := testutil.OverlappingZIP()
	original := bytes.Clone(source)
	r, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	if len(r.File) != 3 {
		t.Fatal("wrong fixture inventory")
	}
	for _, f := range r.File {
		reader, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(reader)
		_ = reader.Close()
		if err != nil || uint64(len(data)) != f.UncompressedSize64 {
			t.Fatalf("fixture payload/CRC failed %s: %v", f.Name, err)
		}
	}
	_, err = OpenPreserved(source, Limits{MaxEntries: 4, MaxTotalBytes: 4096})
	var refusal *Refusal
	if !errors.As(err, &refusal) || !strings.Contains(refusal.Detail, "overlapping archive entries") {
		t.Fatal("did not refuse independent physical overlap", err)
	}
	if !bytes.Equal(source, original) {
		t.Fatal("source altered")
	}
}
