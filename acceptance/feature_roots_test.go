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

// Select only entity values, stylesheet PI and implicit xml prefix; sibling parsing workflows stay planned.
func xmlParsingFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xml", "parsing.feature")
}

// Select named XML editing cases only; sibling workflows stay planned.
func xmlEditingFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xml", "editing.feature")
}

// Only this single sealed canonical outcome is selected for native execution.
func overlapFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "package", "zip-admission.feature")
}

// Select only the exact Unicode QName and byte-offset case from the shared XML names feature.
func xmlNamesFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xml", "names.feature")
}

// Only four exact negative Boolean outcomes are selected; the OPC-like positive stays planned.
func xmlComparisonFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xml", "comparison.feature")
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

// Select only the new-document body/count case, not sibling authoring workflows.
func creationFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "creation.feature")
}

// Select only the direct 10.5pt save-reopen half-point case.
func directFontSizeFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "font-size.feature")
}

// Select only the paragraph getter and body insertion cases, not sibling scenarios.
func paragraphFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "paragraphs.feature")
}

// Select only the named in-memory core-property getter case.
func corePropertiesFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "properties.feature")
}

// Select only the named first-section and background getter case.
func pageLayoutFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "page-layout.feature")
}

// Select named table getter/readback cases, not table editing workflows.
func tableMergeFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "docx", "tables.feature")
}

// Select only direct-range parsing and refusal; sibling formula workflows stay planned.
func formulaReferenceFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xlsx", "formula-references.feature")
}

// The cross-sheet cache case replaces native CACHE-001 one-for-one.
func crossSheetCacheFeaturePath() string {
	return filepath.Join(testutil.ReferenceRoot(), "workflows", "xlsx", "formula-cache.feature")
}

func goCandidateRoot() string {
	return filepath.Join(testutil.ReferenceRoot(), "staging", "go", "behaviors")
}
