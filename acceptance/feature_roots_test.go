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

// This one canonical case replaces native CHAIN-001, without selecting sibling XLSX workflows.
func ownedChainFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xlsx", "calculation-chain-lifecycle.feature")
}

// This exact shared case checks only the in-memory Go run-effect getters.
func runEffectsFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "run-formatting.feature")
}

// Select only the direct in-memory cell merge getter case, not table editing workflows.
func tableMergeFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "tables.feature")
}

// The cross-sheet cache case replaces native CACHE-001 one-for-one.
func crossSheetCacheFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xlsx", "formula-cache.feature")
}

func goCandidateRoot() string {
	return filepath.Join(testutil.ReferenceRoot(), "staging", "go", "behaviors")
}
