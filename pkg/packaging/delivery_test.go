package packaging

import (
	"bytes"
	"errors"
	"io"
	"os"
	"path/filepath"
	"testing"
)

func TestAtomicDeliveryFaultBatch(t *testing.T) {
	fault := errors.New("injected failure")
	for _, phase := range []string{"serialization", "verification", "symlink"} {
		t.Run(phase, func(t *testing.T) {
			dir := t.TempDir()
			dest := filepath.Join(dir, "output")
			original := []byte("original")
			if err := os.WriteFile(dest, original, 0640); err != nil {
				t.Fatal(err)
			}
			write := func(w io.Writer) error { _, err := w.Write([]byte("new")); return err }
			var verify func(string) error
			target := dest
			switch phase {
			case "serialization":
				write = func(w io.Writer) error { _, _ = w.Write([]byte("partial")); return fault }
			case "verification":
				verify = func(string) error { return fault }
			case "symlink":
				target = filepath.Join(dir, "link")
				if err := os.Symlink(dest, target); err != nil {
					t.Fatal(err)
				}
			}
			if err := atomicDeliver(target, write, verify); err == nil {
				t.Fatal("expected failure")
			}
			b, err := os.ReadFile(dest)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(b, original) {
				t.Fatal("destination changed")
			}
			entries, err := os.ReadDir(dir)
			if err != nil {
				t.Fatal(err)
			}
			want := 1
			if phase == "symlink" {
				want = 2
			}
			if len(entries) != want {
				t.Fatalf("temporary files remain: %v", entries)
			}
		})
	}
}
