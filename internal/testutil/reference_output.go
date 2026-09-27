package testutil

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// CheckReferenceOutput refuses output under the shared reference root. Resolve
// existing ancestors so an output directory symlink cannot redirect a write to
// reference assets. It never creates directories or changes reference metadata.
func CheckReferenceOutput(path string) error {
	root, err := resolvedFuturePath(ReferenceRoot())
	if err != nil {
		return err
	}
	output, err := resolvedFuturePath(path)
	if err != nil {
		return err
	}
	rel, err := filepath.Rel(root, output)
	if err != nil {
		return err
	}
	if rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return fmt.Errorf("test output would write shared references: %s", path)
	}
	return nil
}

func resolvedFuturePath(path string) (string, error) {
	path, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	var suffix []string
	for {
		if _, err := os.Lstat(path); err == nil {
			base, err := filepath.EvalSymlinks(path)
			if err != nil {
				return "", err
			}
			for i := len(suffix) - 1; i >= 0; i-- {
				base = filepath.Join(base, suffix[i])
			}
			return base, nil
		} else if !os.IsNotExist(err) {
			return "", err
		}
		parent := filepath.Dir(path)
		if parent == path {
			return "", fmt.Errorf("cannot resolve output ancestor: %s", path)
		}
		suffix = append(suffix, filepath.Base(path))
		path = parent
	}
}
