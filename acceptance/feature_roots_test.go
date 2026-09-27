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

// Only this single sealed canonical outcome is selected for native execution.
func overlapFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "package", "zip-admission.feature")
}

func goCandidateRoot() string {
	return filepath.Join(testutil.ReferenceRoot(), "staging", "go", "behaviors")
}
