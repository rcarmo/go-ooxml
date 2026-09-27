package testutil

import (
	"os"
	"path/filepath"
	"testing"
)

func TestReferencePathsAndOutputGuards(t *testing.T) {
	base := t.TempDir()
	root := filepath.Join(base, "shared")
	if err := os.Mkdir(root, 0755); err != nil {
		t.Fatal(err)
	}
	t.Setenv("OOXML_FIXTURES_ROOT", root)
	alias := filepath.Join(base, "redirect")
	if err := os.Symlink(root, alias); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name, path string
		refuse     bool
	}{
		{"root", root, true},
		{"new descendant", filepath.Join(root, "new", "file"), true},
		{"symlink descendant", filepath.Join(alias, "new", "file"), true},
		{"sibling prefix", root + "-outputs/file", false},
		{"outside", filepath.Join(base, "artifacts", "file"), false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			err := CheckReferenceOutput(tc.path)
			if (err != nil) != tc.refuse {
				t.Fatalf("%s: %v", tc.path, err)
			}
		})
	}
}
