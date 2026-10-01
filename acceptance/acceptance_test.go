package acceptance

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"os"
	"path/filepath"
	"runtime"
	"slices"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// A case is identified by file, stable scenario tag and source line of its
// Examples row (or scenario). Counts alone cannot establish execution coverage.
type caseID struct {
	File string `json:"file"`
	ID   string `json:"id"`
	Line int    `json:"line"`
}
type expectedCase struct {
	caseID
	Name  string `json:"name"`
	Steps int    `json:"steps"`
}
type reportStep struct {
	Name   string `json:"name"`
	Result struct {
		Status string `json:"status"`
	} `json:"result"`
}
type reportElement struct {
	Name string `json:"name"`
	Line int    `json:"line"`
	Type string `json:"type"`
	Tags []struct {
		Name string `json:"name"`
	} `json:"tags"`
	Steps []reportStep `json:"steps"`
}
type reportFeature struct {
	URI      string          `json:"uri"`
	Elements []reportElement `json:"elements"`
}

type world struct {
	source   []byte
	pkg      *packaging.Package
	fixtures map[string]string
}

func (w *world) fixture(name string) error {
	if !filepath.IsLocal(name) {
		return fmt.Errorf("unsafe fixture %q", name)
	}
	data, err := os.ReadFile(testutil.FixturePath(name))
	if err != nil {
		return err
	}
	w.source = data
	sum := sha256.Sum256(data)
	w.fixtures[name] = hex.EncodeToString(sum[:])
	return nil
}
func (w *world) open() error { var err error; w.pkg, err = packaging.OpenBytes(w.source); return err }
func (w *world) contains(name string) error {
	if w.pkg == nil {
		return fmt.Errorf("package not open")
	}
	p, err := w.pkg.GetPart(name)
	if err != nil {
		return err
	}
	content, err := p.Content()
	if err != nil {
		return err
	}
	if len(content) == 0 {
		return fmt.Errorf("empty part %s", name)
	}
	return nil
}

func TestAcceptance(t *testing.T) {
	// The selected feature must belong to a verified release or explicit clean
	// candidate checkout, even when this test is run without the rest of the suite.
	loadReferencePin(t)
	w := &world{fixtures: map[string]string{}}
	readbackDir := t.TempDir()
	initializer := func(sc *godog.ScenarioContext) {
		tableReadback := &tableTextReadbackState{}
		safetySteps(sc)
		zip64Steps(sc)
		overlapSteps(sc)
		bzipAdmissionSteps(sc)
		negativeBudgetSteps(sc)
		xmlComparisonSteps(sc)
		unicodeQNameSteps(sc)
		if xmlLexicalCandidate() || postBatch2Reference() {
			lexicalXMLSteps(sc)
		}
		immutableLeafSteps(sc)
		attributeSpliceSteps(sc)
		attributeRefusalSteps(sc)
		elementRemovalSteps(sc)
		elementRemovalRefusalSteps(sc)
		elementReplacementCustodySteps(sc)
		elementReplacementRefusalSteps(sc)
		childInsertionCustodySteps(sc)
		childInsertionRefusalSteps(sc)
		childNamespaceMatrixSteps(sc)
		xmlSafety := &xmlSafetyState{}
		xmlEntityValuesSteps(sc, xmlSafety)
		xmlSafetySteps(sc, xmlSafety)
		xmlWhitespaceRoundtripSteps(sc)
		profileReasons := &packageReasonState{}
		descriptorCollisionSteps(sc, profileReasons)
		limitSteps(sc)
		preservedSteps(sc)
		opcPreserveUnrelatedSteps(sc)
		opcCorpusNoopSteps(sc)
		opcDetachedByteSteps(sc)
		packageReasonSteps(sc, profileReasons)
		receiptSteps(sc)
		xmlSteps(sc)
		graphSteps(sc)
		graphEditSteps(sc)
		graphDeleteSteps(sc)
		graphTargetFormSteps(sc)
		wordSteps(sc)
		slideSteps(sc)
		notesEditSteps(sc)
		notesMultilineSteps(sc)
		cellSteps(sc)
		vocabularySteps(sc)
		storySteps(sc)
		spanSteps(sc)
		replaceSpanSteps(sc)
		spanBatchSteps(sc)
		searchPolicySteps(sc)
		attributeSteps(sc)
		whitespaceSteps(sc)
		replaceAllSteps(sc)
		searchScopeSteps(sc)
		runBoundarySteps(sc)
		spanGuardSteps(sc)
		styleIndexSteps(sc)
		xmlNormalizationSteps(sc)
		xmlInsertSteps(sc)
		subtreeSteps(sc)
		xmlConformanceSteps(sc)
		officeLinkSteps(sc)
		formulaSteps(sc)
		canonicalFormulaAnalysisSteps(sc)
		remainingFormulaSteps(sc)
		directRangeSteps(sc)
		cacheSteps(sc)
		canonicalCacheSteps(sc)
		runEffectsSteps(sc)
		runVerticalAlignSteps(sc)
		tableMergeGetterSteps(sc)
		tableValueSteps(sc, tableReadback)
		documentCreationSteps(sc, tableReadback)
		documentPropertyGetterSteps(sc, tableReadback)
		paragraphTextGetterSteps(sc)
		bodyInsertOrderSteps(sc)
		tableTextReadbackSteps(sc, tableReadback)
		runFormattingReadbackSteps(sc, readbackDir, tableReadback)
		directFontSizeSteps(sc, readbackDir)
		imageReplaceSteps(sc)
		remapSteps(sc)
		commentMIMESteps(sc)
		if pptxManipulationCandidate() {
			pptxManipulationSteps(sc)
		}
		if formattingCandidate() {
			formattingSteps(sc)
		}
		sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
			w.source = nil
			w.pkg = nil
			return ctx, nil
		})
		sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
			if w.pkg != nil {
				_ = w.pkg.Close()
			}
			return ctx, nil
		})
		sc.Step(`^the repository fixture "([^"]+)"$`, w.fixture)
		sc.Step(`^I open its Office package$`, w.open)
		sc.Step(`^the package contains a nonempty part "([^"]+)"$`, w.contains)
	}
	expected, inventory, err := inventoryCases()
	if err != nil {
		t.Fatal(err)
	}
	if len(expected) == 0 {
		t.Fatal("no implemented cases inventoried")
	}
	if pptxManipulationCandidate() {
		pptxExpected, pptxInventory, err := pptxManipulationInventory()
		if err != nil {
			t.Fatal(err)
		}
		for key, item := range pptxExpected {
			if _, exists := expected[key]; exists {
				t.Fatalf("PPTX case conflicts with earlier selection: %+v", key)
			}
			expected[key] = item
		}
		inventory = append(inventory, pptxInventory...)
	}
	if formattingCandidate() {
		more, entries, err := formattingInventory()
		if err != nil {
			t.Fatal(err)
		}
		for key, item := range more {
			if _, exists := expected[key]; exists {
				t.Fatalf("formatting case conflicts: %+v", key)
			}
			expected[key] = item
		}
		inventory = append(inventory, entries...)
	}
	if batch2Candidate() {
		batchExpected, batchInventory, err := batch2Inventory()
		if err != nil {
			t.Fatal(err)
		}
		for key, item := range batchExpected {
			if previous, exists := expected[key]; exists {
				if previous != item {
					t.Fatalf("batch-2 case conflicts with earlier selection: %+v", key)
				}
				continue
			}
			expected[key] = item
		}
		for _, row := range batchInventory {
			key := row["case"].(caseID)
			if key.ID == opcPreserveUnrelatedCaseID || key.ID == opcCorpusNoopCaseID || key.ID == opcDetachedByteCaseID {
				continue
			}
			inventory = append(inventory, row)
		}
	}
	var combined []reportFeature
	selections := []struct {
		name, path, tags string
	}{
		{"go-ooxml-native", goFeatureRoot(), "@implemented && @go && ~@CHAIN-001 && ~@CACHE-001"},
		{"go-ooxml-overlap", overlapFeaturePath(), overlapCaseID},
		{"go-ooxml-bzip-admission", overlapFeaturePath(), bzipAdmissionCaseID},
		{"go-ooxml-unsafe-members", overlapFeaturePath(), unsafeMembersCaseID},
		{"go-ooxml-opc-preserve-unrelated", packagePreservationFeaturePath(), opcPreserveUnrelatedCaseID},
		{"go-ooxml-opc-corpus-noop", packagePreservationFeaturePath(), opcCorpusNoopCaseID},
		{"go-ooxml-opc-detached-bytes", packagePreservationFeaturePath(), opcDetachedByteCaseID},
		{"go-ooxml-negative-budget", negativeBudgetFeaturePath(), negativeBudgetCaseID},
		{"go-ooxml-unicode-qname", xmlNamesFeaturePath(), unicodeQNameCaseID},
		{"go-ooxml-immutable-leaf", xmlEditingFeaturePath(), immutableLeafCaseID},
		{"go-ooxml-attribute-splice", xmlEditingFeaturePath(), attributeSpliceCaseID},
		{"go-ooxml-attribute-refusal", xmlEditingFeaturePath(), attributeRefusalCaseID},
		{"go-ooxml-element-removal", xmlEditingFeaturePath(), elementRemovalCaseID},
		{"go-ooxml-element-removal-refusal", xmlEditingFeaturePath(), elementRemovalRefusalCaseID},
		{"go-ooxml-element-replacement-custody", xmlEditingFeaturePath(), elementReplacementCustodyCaseID},
		{"go-ooxml-element-replacement-refusal", xmlEditingFeaturePath(), elementReplacementRefusalCaseID},
		{"go-ooxml-child-insertion-custody", xmlEditingFeaturePath(), childInsertionCustodyCaseID},
		{"go-ooxml-child-insertion-refusal", xmlEditingFeaturePath(), childInsertionRefusalCaseID},
		{"go-ooxml-child-namespace-matrix", xmlEditingFeaturePath(), childNamespaceMatrixCaseID},
		{"go-ooxml-xml-entity-values", xmlParsingFeaturePath(), xmlEntityValuesCaseID},
		{"go-ooxml-xml-whitespace-roundtrip", xmlParsingFeaturePath(), xmlWhitespaceCaseID},
		{"go-ooxml-xml-stylesheet-pi", xmlParsingFeaturePath(), xmlStylesheetPICaseID},
		{"go-ooxml-implicit-xml-prefix", xmlParsingFeaturePath(), xmlImplicitPrefixCaseID},
		{"go-ooxml-expanded-attribute-lookup", xmlParsingFeaturePath(), xmlExpandedAttributeCaseID},
		{"go-ooxml-xml-significant", xmlComparisonFeaturePath(), xmlSignificantCaseID},
		{"go-ooxml-xml-prefix-binding", xmlComparisonFeaturePath(), xmlPrefixBindingCaseID},
		{"go-ooxml-xml-unsafe", xmlComparisonFeaturePath(), xmlUnsafeCaseID},
		{"go-ooxml-xml-markup", xmlComparisonFeaturePath(), xmlMarkupCaseID},
		{"go-ooxml-descriptor-collision", descriptorIntegrityFeaturePath(), descriptorCollisionCaseID},
		{"go-ooxml-owned-chain", ownedChainFeaturePath(), ownedChainCaseID},
		{"go-ooxml-cross-sheet-cache", crossSheetCacheFeaturePath(), crossSheetCacheCaseID},
		{"go-ooxml-direct-range-parsing", formulaReferenceFeaturePath(), directRangeParsingCaseID},
		{"go-ooxml-direct-range-refusal", formulaReferenceFeaturePath(), directRangeRefusalCaseID},
		{"go-ooxml-formula-analysis-counts", formulaReferenceFeaturePath(), formulaAnalysisCountsCaseID},
		{"go-ooxml-formula-quoted-sheet-flags", formulaReferenceFeaturePath(), formulaQuotedSheetFlagsCaseID},
		{"go-ooxml-formula-analysis-refusal", formulaReferenceFeaturePath(), formulaAnalysisRefusalCaseID},
		{"go-ooxml-formula-literal-punctuation", formulaReferenceFeaturePath(), formulaLiteralPunctuationCaseID},
		{"go-ooxml-static-remap-exact", formulaReferenceFeaturePath(), staticRemapExactCaseID},
		{"go-ooxml-static-remap-refusal", formulaReferenceFeaturePath(), staticRemapRefusalCaseID},
		{"go-ooxml-static-reference-properties", formulaReferenceFeaturePath(), staticReferencePropertiesCaseID},
		{"go-ooxml-run-effects", runEffectsFeaturePath(), runEffectsCaseID},
		{"go-ooxml-run-underline-style", runEffectsFeaturePath(), runUnderlineCaseID},
		{"go-ooxml-run-font-name", runEffectsFeaturePath(), runFontNameCaseID},
		{"go-ooxml-run-color-getter", runEffectsFeaturePath(), runColorCaseID},
		{"go-ooxml-run-highlight", runEffectsFeaturePath(), runHighlightCaseID},
		{"go-ooxml-run-vertical-align", runEffectsFeaturePath(), runVerticalAlignCaseID},
		{"go-ooxml-roundtrip-selected-formatting", runEffectsFeaturePath(), runRoundtripFormattingCaseID},
		{"go-ooxml-table-merge-properties", tableMergeFeaturePath(), tableMergeCaseID},
		{"go-ooxml-table-dimensions", tableMergeFeaturePath(), tableDimensionsCaseID},
		{"go-ooxml-table-cell-access", tableMergeFeaturePath(), tableCellAccessCaseID},
		{"go-ooxml-table-cell-text", tableMergeFeaturePath(), tableCellTextCaseID},
		{"go-ooxml-table-row-counts", tableMergeFeaturePath(), tableRowCountsCaseID},
		{"go-ooxml-new-empty-body", creationFeaturePath(), newEmptyBodyCaseID},
		{"go-ooxml-roundtrip-table-text", tableMergeFeaturePath(), roundtripTableTextCaseID},
		{"go-ooxml-core-properties-getters", corePropertiesFeaturePath(), corePropertiesCaseID},
		{"go-ooxml-section-title-background", pageLayoutFeaturePath(), sectionTitleBackgroundCaseID},
		{"go-ooxml-paragraph-text-getter", paragraphFeaturePath(), paragraphTextGetterCaseID},
		{"go-ooxml-paragraph-alignment", paragraphFeaturePath(), paragraphAlignmentCaseID},
		{"go-ooxml-paragraph-spacing", paragraphFeaturePath(), paragraphSpacingCaseID},
		{"go-ooxml-paragraph-toggles", paragraphFeaturePath(), paragraphTogglesCaseID},
		{"go-ooxml-paragraph-runs", paragraphFeaturePath(), paragraphRunsCaseID},
		{"go-ooxml-body-insert-order", paragraphFeaturePath(), bodyInsertOrderCaseID},
		{"go-ooxml-direct-font-size-half-points", directFontSizeFeaturePath(), directFontSizeCaseID},
	}
	if pptxManipulationCandidate() {
		records, err := pptxManipulationRecords()
		if err != nil {
			t.Fatal(err)
		}
		for id, record := range records {
			selections = append(selections, struct{ name, path, tags string }{"go-pptx-manipulation-" + strings.TrimPrefix(id, "@"), testutil.ReferencePath(filepath.FromSlash(record.Feature)), id})
		}
	}
	if formattingCandidate() {
		records, err := formattingRecords()
		if err != nil {
			t.Fatal(err)
		}
		for id, record := range records {
			selections = append(selections, struct{ name, path, tags string }{"go-pptx-formatting-" + strings.TrimPrefix(id, "@"), testutil.ReferencePath(filepath.FromSlash(record.Feature)), id})
		}
	}
	if batch2Candidate() {
		for relative, ids := range batch2Paths {
			for _, id := range ids {
				if id == opcPreserveUnrelatedCaseID || id == opcCorpusNoopCaseID || id == opcDetachedByteCaseID {
					continue
				}
				selections = append(selections, struct{ name, path, tags string }{"go-batch2-" + strings.TrimPrefix(id, "@"), testutil.ReferencePath(filepath.FromSlash(relative)), id})
			}
		}
	}
	if xmlLexicalCandidate() || postBatch2Reference() {
		for _, item := range []struct{ name, path, id string }{
			{"xml-offsets", xmlParsingFeaturePath(), xmlParseOffsetsID}, {"xml-line-endings", xmlParsingFeaturePath(), xmlLineEndingsID},
			{"xml-refusals", xmlParsingFeaturePath(), xmlParseRefusalsID}, {"xml-bounds", xmlParsingFeaturePath(), xmlParseBoundsID},
			{"xml-invalid-qname", xmlNamesFeaturePath(), xmlInvalidQNameID}, {"xml-nbsp", xmlNamesFeaturePath(), xmlOutsideNBSPID},
			{"xml-escape-values", xmlParsingFeaturePath(), xmlEscapingValuesID}, {"xml-escape-invalid", xmlParsingFeaturePath(), xmlEscapingInvalidID},
			{"xml-apply-edits", xmlEditingFeaturePath(), xmlApplyEditsID},
		} {
			selections = append(selections, struct{ name, path, tags string }{"go-ooxml-" + item.name, item.path, item.id})
		}
	}
	if xmlSafetyCandidate() {
		for _, item := range []struct{ name, id string }{
			{"go-ooxml-xml-prototype-safe-attributes", xmlPrototypeCaseID},
			{"go-ooxml-xml-immutable-namespace-metadata", xmlNamespaceCaseID},
			{"go-ooxml-xml-typed-parse-error", xmlMalformedCaseID},
		} {
			selections = append(selections, struct{ name, path, tags string }{item.name, xmlParsingFeaturePath(), item.id})
		}
	}
	if packageReasonsCandidate() {
		for _, item := range []struct{ name, path, id string }{
			{"go-opc-open-reasons", packagePreservationFeaturePath(), opcReasonOpenID},
			{"go-opc-save-missing-target", packagePreservationFeaturePath(), opcReasonSaveID},
			{"go-opc-save-symlink", packagePreservationFeaturePath(), opcReasonSymlinkID},
			{"go-zip32-reader-reasons", zip32FeaturePath(), zipReasonReaderID},
			{"go-zip32-writer-reasons", zip32FeaturePath(), zipReasonWriterID},
			{"go-zip32-bounds-reasons", zip32FeaturePath(), zipReasonBoundsID},
		} {
			selections = append(selections, struct{ name, path, tags string }{item.name, item.path, item.id})
		}
	}
	if portableTransactionCandidate() {
		selections = append(selections,
			struct{ name, path, tags string }{"go-opc-deferred-transaction", packagePreservationFeaturePath(), opcDeferredTransactionID},
			struct{ name, path, tags string }{"go-opc-opaque-transaction", packagePreservationFeaturePath(), opcOpaqueTransactionID},
		)
	}
	for _, selection := range selections {
		var output bytes.Buffer
		steps := initializer
		if postBatch2Reference() && selection.tags == opcPreserveUnrelatedCaseID {
			steps = batch2Steps
		}
		if strings.HasPrefix(selection.name, "go-batch2-") {
			steps = batch2Steps
		}
		suite := godog.TestSuite{Name: selection.name, Options: &godog.Options{Format: "cucumber", Output: &output, Paths: []string{selection.path}, Tags: selection.tags, Strict: true, Concurrency: 1}, ScenarioInitializer: steps}
		code := suite.Run()
		var features []reportFeature
		if err := json.Unmarshal(output.Bytes(), &features); err != nil {
			t.Fatalf("%s invalid report: %v (%s)", selection.name, err, output.String())
		}
		if code != 0 {
			t.Errorf("%s Godog failed (%d): %s", selection.name, code, output.String())
		}
		combined = append(combined, features...)
	}
	data, err := json.Marshal(combined)
	if err != nil {
		t.Fatal(err)
	}
	dir := os.Getenv("OOXML_REPORT_DIR")
	if dir == "" {
		dir = "../reports/acceptance"
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "cucumber.json"), data, 0644); err != nil {
		t.Fatal(err)
	}
	writeJSON(t, filepath.Join(dir, "inventory.json"), inventory)
	scope := "implemented native except CHAIN-001 and CACHE-001, each replaced one-for-one by its exact canonical case, plus overlap, negative-budget, descriptor-collision, four exact negative XML comparison IDs, one Unicode QName/offset case, one immutable XML leaf seed, one exact 100-choice child-namespace matrix, one three-step public XML whitespace roundtrip, three exact attribute-splice rows, one exact duplicate-attribute refusal, one exact two-target element removal and selected in-memory run-formatting and paragraph getter cases plus one direct font-size save-reopen case, 13 direct-range and 32 static formula-analysis/remap API cases"
	if pptxManipulationCandidate() {
		scope += "; explicit sealed PPTX manipulation candidate: 20 selected cases/179 steps (within aggregate 498 cases/1924 steps), not inherited batch-2 execution credit"
	}
	if formattingCandidate() {
		scope += "; explicit sealed PPTX formatting candidate: 20 selected cases/182 steps, separate from prior manipulation 20/179; aggregate reconciled by exact case IDs"
	}
	writeJSON(t, filepath.Join(dir, "environment.json"), map[string]any{"go": runtime.Version(), "fixture_sha256": w.fixtures, "scope": scope, "external_executed": false})
	if err := reconcile(expected, data); err != nil {
		t.Error(err)
	}
}

func writeJSON(t *testing.T, path string, value any) {
	t.Helper()
	b, err := json.MarshalIndent(value, "", "  ")
	if err != nil {
		t.Fatal(err)
	}
	if err = os.WriteFile(path, append(b, '\n'), 0644); err != nil {
		t.Fatal(err)
	}
}
func stableID(tags []string) (string, error) {
	id := ""
	for _, tag := range tags {
		if nativeIDPattern.MatchString(tag) || (batch2Candidate() && batch2SelectedID(tag)) || (pptxManipulationCandidate() && pptxManipulationSelectedID(tag)) || (formattingCandidate() && strings.HasPrefix(tag, "@id-pptx-formatting-")) || tag == overlapCaseID || tag == bzipAdmissionCaseID || tag == unsafeMembersCaseID || tag == opcPreserveUnrelatedCaseID || tag == opcCorpusNoopCaseID || tag == opcDetachedByteCaseID || tag == opcReasonOpenID || tag == opcReasonSaveID || tag == opcReasonSymlinkID || tag == opcDeferredTransactionID || tag == opcOpaqueTransactionID || tag == zipReasonReaderID || tag == zipReasonWriterID || tag == zipReasonBoundsID || tag == immutableLeafCaseID || tag == attributeSpliceCaseID || tag == attributeRefusalCaseID || tag == elementRemovalCaseID || tag == elementRemovalRefusalCaseID || tag == elementReplacementCustodyCaseID || tag == elementReplacementRefusalCaseID || tag == childInsertionCustodyCaseID || tag == childInsertionRefusalCaseID || tag == childNamespaceMatrixCaseID || tag == xmlEntityValuesCaseID || tag == xmlPrototypeCaseID || tag == xmlNamespaceCaseID || tag == xmlMalformedCaseID || tag == xmlWhitespaceCaseID || tag == xmlStylesheetPICaseID || tag == xmlImplicitPrefixCaseID || tag == xmlExpandedAttributeCaseID || tag == unicodeQNameCaseID || ((xmlLexicalCandidate() || postBatch2Reference()) && slices.Contains(xmlLexicalIDs, tag)) || tag == xmlSignificantCaseID || tag == xmlPrefixBindingCaseID || tag == xmlUnsafeCaseID || tag == xmlMarkupCaseID || tag == negativeBudgetCaseID || tag == descriptorCollisionCaseID || tag == ownedChainCaseID || tag == crossSheetCacheCaseID || tag == directRangeParsingCaseID || tag == directRangeRefusalCaseID || tag == formulaAnalysisCountsCaseID || tag == formulaQuotedSheetFlagsCaseID || tag == formulaAnalysisRefusalCaseID || tag == formulaLiteralPunctuationCaseID || tag == staticRemapExactCaseID || tag == staticRemapRefusalCaseID || tag == staticReferencePropertiesCaseID || tag == runEffectsCaseID || tag == runUnderlineCaseID || tag == runFontNameCaseID || tag == runColorCaseID || tag == runHighlightCaseID || tag == runVerticalAlignCaseID || tag == runRoundtripFormattingCaseID || tag == tableMergeCaseID || tag == tableDimensionsCaseID || tag == tableCellAccessCaseID || tag == tableCellTextCaseID || tag == tableRowCountsCaseID || tag == newEmptyBodyCaseID || tag == roundtripTableTextCaseID || tag == corePropertiesCaseID || tag == sectionTitleBackgroundCaseID || tag == paragraphTextGetterCaseID || tag == paragraphAlignmentCaseID || tag == paragraphSpacingCaseID || tag == paragraphTogglesCaseID || tag == paragraphRunsCaseID || tag == bodyInsertOrderCaseID || tag == directFontSizeCaseID {
			if id != "" {
				return "", fmt.Errorf("multiple IDs: %v", tags)
			}
			id = tag
		}
	}
	if id == "" {
		return "", fmt.Errorf("missing stable ID: %v", tags)
	}
	return id, nil
}

func reconcile(expected map[caseID]expectedCase, data []byte) error {
	var features []reportFeature
	if err := json.Unmarshal(data, &features); err != nil {
		return err
	}
	seen := map[caseID]bool{}
	for _, f := range features {
		for _, e := range f.Elements {
			if e.Type != "scenario" {
				return fmt.Errorf("unexpected report element %q", e.Type)
			}
			tags := []string{}
			for _, tag := range e.Tags {
				tags = append(tags, tag.Name)
			}
			id, err := stableID(tags)
			if err != nil {
				return err
			}
			key := caseID{filepath.ToSlash(filepath.Clean(f.URI)), id, e.Line}
			want, ok := expected[key]
			if !ok {
				return fmt.Errorf("unplanned result %+v", key)
			}
			if seen[key] {
				return fmt.Errorf("duplicate result %+v", key)
			}
			seen[key] = true
			if e.Name != want.Name || len(e.Steps) != want.Steps {
				return fmt.Errorf("name/steps mismatch %+v", key)
			}
			for _, step := range e.Steps {
				if step.Result.Status != "passed" {
					return fmt.Errorf("%+v step %q: %s", key, step.Name, step.Result.Status)
				}
			}
		}
	}
	for key := range expected {
		if !seen[key] {
			return fmt.Errorf("missing result %+v", key)
		}
	}
	return nil
}
