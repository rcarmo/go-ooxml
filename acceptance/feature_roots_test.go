package acceptance

import (
	"path/filepath"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// Both the runner and inventory must use the same reference checkout. The
// candidate override is resolved by ReferenceRoot, not by looking beside the
// Go source tree when a different sealed distribution is under test.
func goFeatureRoot() string {
	return filepath.Join(testutil.ReferenceRoot(), "staging", "go", "features")
}

func goCandidateRoot() string {
	return filepath.Join(testutil.ReferenceRoot(), "staging", "go", "behaviors")
}
