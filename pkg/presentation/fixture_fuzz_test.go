package presentation

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func FuzzPresentationFixtureMutation(f *testing.F) {
	fixtures := []string{
		"bullet_points.pptx",
		"comments.pptx",
		"hidden_slides.pptx",
		"images.pptx",
		"layouts.pptx",
		"minimal.pptx",
		"multiple_masters.pptx",
		"notes.pptx",
		"shapes.pptx",
		"tables.pptx",
		"title_slide.pptx",
	}

	for _, name := range fixtures {
		data, err := os.ReadFile(fixturePath(name))
		if err != nil {
			continue
		}
		f.Add(data, uint16(0), byte(0))
	}

	f.Fuzz(func(t *testing.T, data []byte, offset uint16, xor byte) {
		if len(data) == 0 || len(data) > 8<<20 {
			return
		}

		mutated := testutil.MutateBytes(data, offset, xor)
		pres, err := OpenReader(bytes.NewReader(mutated), int64(len(mutated)))
		if err != nil {
			return
		}
		defer pres.Close()

		_ = pres.Slides()
		_ = pres.Layouts()

		path := filepath.Join(t.TempDir(), "fixture.pptx")
		if err := pres.SaveAs(path); err != nil {
			return
		}

		reopen, err := Open(path)
		if err != nil {
			return
		}
		defer reopen.Close()

		_ = reopen.Slides()
		_ = reopen.Layouts()
	})
}

func fixturePath(name string) string {
	return filepath.Join("..", "..", "testdata", "pptx", name)
}
