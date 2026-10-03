package acceptance

import (
	"fmt"
	"os"
	"path/filepath"
	"reflect"
)

// contract20SaveFaults tests actual SaveAs destinations: an existing symlink
// to a regular sentinel, an absent file whose parent does not exist, and a
// directory. No test-only injection or unrelated "prior destination" counts.
func contract20SaveFaults(temp string, save func(string) error) error {
	const sentinel = "existing destination sentinel"
	if err := os.MkdirAll(temp, 0700); err != nil {
		return err
	}
	target := filepath.Join(temp, "existing-target.xlsx")
	link := filepath.Join(temp, "existing-link.xlsx")
	if err := os.WriteFile(target, []byte(sentinel), 0600); err != nil {
		return err
	}
	if err := os.Symlink(filepath.Base(target), link); err != nil {
		return err
	}
	absentParent := filepath.Join(temp, "absent-parent")
	absent := filepath.Join(absentParent, "new.xlsx")
	directory := filepath.Join(temp, "directory-destination")
	if err := os.Mkdir(directory, 0700); err != nil {
		return err
	}
	list := func() ([]string, error) {
		entries, err := os.ReadDir(temp)
		if err != nil {
			return nil, err
		}
		names := make([]string, 0, len(entries))
		for _, entry := range entries {
			names = append(names, entry.Name())
		}
		return names, nil
	}
	prior, err := list()
	if err != nil {
		return err
	}
	for _, dest := range []string{link, absent, directory} {
		if err := save(dest); err == nil {
			return fmt.Errorf("SaveAs unexpectedly accepted %s", dest)
		}
		after, err := list()
		if err != nil || !reflect.DeepEqual(prior, after) {
			return fmt.Errorf("failed SaveAs changed temporary directory %v -> %v: %v", prior, after, err)
		}
	}
	if targetName, err := os.Readlink(link); err != nil || targetName != filepath.Base(target) {
		return fmt.Errorf("existing symlink target changed: %q/%v", targetName, err)
	}
	body, err := os.ReadFile(target)
	if err != nil || string(body) != sentinel {
		return fmt.Errorf("existing target content changed: %q/%v", body, err)
	}
	if _, err := os.Lstat(absentParent); !os.IsNotExist(err) {
		return fmt.Errorf("absent destination parent created: %v", err)
	}
	if info, err := os.Lstat(directory); err != nil || !info.IsDir() {
		return fmt.Errorf("directory target changed: %v", err)
	}
	return nil
}
