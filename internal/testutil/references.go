package testutil

import (
	"os"
	"path/filepath"
	"runtime"
)

// ReferenceRoot selects the shared read-only reference checkout. The explicit
// override supports testing a candidate distribution before the submodule tag is
// sealed. Missing inputs must fail at the caller; there is no legacy fallback.
func ReferenceRoot() string {
	if root := os.Getenv("OOXML_FIXTURES_ROOT"); root != "" {
		return root
	}
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		panic("cannot locate shared fixture checkout")
	}
	return filepath.Join(filepath.Dir(file), "..", "..", "references", "fixtures-ooxml")
}

// ReferencePath locates a distribution input. Callers supply trusted relative
// constants, never output paths or untrusted archive-member names.
func ReferencePath(parts ...string) string {
	return filepath.Join(append([]string{ReferenceRoot()}, parts...)...)
}

// FixturePath locates the current Go corpus, including frozen generated inputs.
// New test outputs belong under t.TempDir or consumer-local artifacts instead.
func FixturePath(parts ...string) string {
	return ReferencePath(append([]string{"fixtures", "go-current", "testdata"}, parts...)...)
}
