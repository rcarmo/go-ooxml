package packaging

import (
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

// atomicDeliver guarantees destination preservation on failures before rename.
// It syncs the temporary file; directory fsync/crash durability is not promised.
func atomicDeliver(filePath string, write func(io.Writer) error, verify func(string) error) error {
	if filePath == "" {
		return utils.ErrPathNotSet
	}
	clean := filepath.Clean(filePath)
	mode := os.FileMode(0600)
	// Reject symlinks rather than changing between follow-on-write and replace-
	// link semantics. Applications should choose the real destination explicitly.
	if info, err := os.Lstat(clean); err == nil {
		if !info.Mode().IsRegular() {
			return fmt.Errorf("destination is not a regular file: %s", clean)
		}
		mode = info.Mode().Perm()
	} else if !os.IsNotExist(err) {
		return err
	}
	f, err := os.CreateTemp(filepath.Dir(clean), ".ooxml-*")
	if err != nil {
		return err
	}
	temporary := f.Name()
	defer func() { _ = f.Close(); _ = os.Remove(temporary) }()
	if err = f.Chmod(mode); err != nil {
		return err
	}
	if err = write(f); err != nil {
		return err
	}
	if err = f.Sync(); err != nil {
		return err
	}
	if err = f.Close(); err != nil {
		return err
	}
	if verify != nil {
		if err = verify(temporary); err != nil {
			return err
		}
	}
	return os.Rename(temporary, clean)
}
