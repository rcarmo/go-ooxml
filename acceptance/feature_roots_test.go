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

// Only the negative-budget scenario outline in this canonical feature is selected.
func negativeBudgetFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "package", "admission-limit-configuration.feature")
}

// The unsigned-descriptor collision is selected without enabling sibling workflows.
func descriptorIntegrityFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "package", "data-descriptor-integrity.feature")
}

func goCandidateRoot() string {
	return filepath.Join(testutil.ReferenceRoot(), "staging", "go", "behaviors")
}
