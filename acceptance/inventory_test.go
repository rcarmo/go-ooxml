package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"slices"
	"strings"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
)

// Inventory all feature files, including planned and external contracts. Native
// execution selects only implemented cases; excluded cases remain in the map.
func inventoryCases() (map[caseID]expectedCase, []map[string]any, error) {
	expected := map[caseID]expectedCase{}
	var inventory []map[string]any
	seen := map[string]string{}
	counter := 0
	next := func() string { counter++; return fmt.Sprint(counter) }
	err := filepath.WalkDir(goFeatureRoot(), func(path string, d os.DirEntry, err error) error {
		return inventoryFeature(path, d, err, expected, &inventory, seen, next, false)
	})
	if err != nil {
		return nil, nil, err
	}
	paths := []string{overlapFeaturePath(), packagePreservationFeaturePath(), xmlComparisonFeaturePath(), xmlNamesFeaturePath(), xmlEditingFeaturePath(), xmlParsingFeaturePath(), negativeBudgetFeaturePath(), retainedVisibilityFeaturePath(), descriptorIntegrityFeaturePath(), ownedChainFeaturePath(), crossSheetCacheFeaturePath(), formulaReferenceFeaturePath(), runEffectsFeaturePath(), tableMergeFeaturePath(), creationFeaturePath(), corePropertiesFeaturePath(), pageLayoutFeaturePath(), paragraphFeaturePath(), directFontSizeFeaturePath()}
	if packageReasonsCandidate() {
		paths = append(paths, zip32FeaturePath())
	}
	for _, path := range paths {
		if uniformAPI18Candidate() && path == formulaReferenceFeaturePath() {
			continue // All nine formula IDs are compiled against the stronger ledger below.
		}
		if err = inventoryFeature(path, nil, nil, expected, &inventory, seen, next, true); err != nil {
			return nil, nil, err
		}
	}
	if uniformAPI18Candidate() {
		more, entries, err := uniformInventory()
		if err != nil {
			return nil, nil, err
		}
		for key, item := range more {
			if _, exists := expected[key]; exists {
				return nil, nil, fmt.Errorf("uniform API case conflicts with earlier selection: %+v", key)
			}
			expected[key] = item
		}
		inventory = append(inventory, entries...)
	}
	return expected, inventory, nil
}

var nativeIDPattern = regexp.MustCompile(`^@[A-Z]+-[0-9]{3}$`)

func hasScenarioTag(s *messages.Scenario, id string) bool {
	for _, tag := range s.Tags {
		if tag.Name == id {
			return true
		}
	}
	return false
}

const immutableLeafCaseID = "@id-xml-go-immutable-leaf-seed"
const attributeSpliceCaseID = "@id-xml-go-attribute-splice-custody"
const attributeRefusalCaseID = "@id-xml-go-attribute-batch-refusal"
const elementRemovalCaseID = "@id-xml-go-element-removal-custody"
const elementRemovalRefusalCaseID = "@id-xml-go-element-removal-refusal"
const childInsertionCustodyCaseID = "@id-xml-go-child-insertion-custody"
const childInsertionRefusalCaseID = "@id-xml-go-child-insertion-refusal"
const unicodeQNameCaseID = "@id-xml-unicode-qname-components"
const xmlSignificantCaseID = "@id-xml-comparison-significant-content"
const xmlPrefixBindingCaseID = "@id-xml-comparison-prefix-attribute-binding"
const xmlUnsafeCaseID = "@id-xml-comparison-unsafe-input"
const xmlMarkupCaseID = "@id-xml-comparison-processing-instructions-and-comments"

const overlapCaseID = "@id-zip-physical-member-overlap-refusal"
const negativeBudgetCaseID = "@id-package-admission-negative-budget"
const retainedVisibilityEvidenceCaseID = "@id-pptx-slide-visibility-retained-inputs"
const descriptorCollisionCaseID = "@id-zip-unsigned-descriptor-signature-collision"
const ownedChainCaseID = "@id-xlsx-owned-calculation-chain-invalidation"
const retiredChainCaseID = "@CHAIN-001"
const crossSheetCacheCaseID = "@id-xlsx-cross-sheet-cache-invalidation"
const runEffectsCaseID = "@id-docx-go-run-effects-getters"
const runUnderlineCaseID = "@id-docx-go-run-underline-style"
const runFontNameCaseID = "@id-docx-go-run-font-name"
const runColorCaseID = "@id-docx-go-run-color-getter"
const runHighlightCaseID = "@id-docx-go-run-highlight"
const runVerticalAlignCaseID = "@id-docx-go-run-vertical-align"
const runRoundtripFormattingCaseID = "@id-docx-go-roundtrip-selected-formatting"
const tableMergeCaseID = "@id-docx-go-table-merge-properties"
const tableDimensionsCaseID = "@id-docx-go-table-dimensions-getters"
const tableCellAccessCaseID = "@id-docx-go-table-cell-access"
const tableCellTextCaseID = "@id-docx-go-table-cell-text-getters"
const tableRowCountsCaseID = "@id-docx-go-table-row-counts"
const newEmptyBodyCaseID = "@id-docx-go-new-empty-body"
const roundtripTableTextCaseID = "@id-docx-go-roundtrip-table-text"
const corePropertiesCaseID = "@id-docx-go-core-properties-getters"
const sectionTitleBackgroundCaseID = "@id-docx-go-section-title-background-getters"
const paragraphTextGetterCaseID = "@id-docx-go-paragraph-text-getter"
const paragraphAlignmentCaseID = "@id-docx-go-paragraph-alignment-getter"
const paragraphSpacingCaseID = "@id-docx-go-paragraph-spacing-getters"
const paragraphTogglesCaseID = "@id-docx-go-paragraph-advanced-toggles"
const paragraphRunsCaseID = "@id-docx-go-paragraph-multiple-runs"
const bodyInsertOrderCaseID = "@id-docx-go-body-insert-order"
const directFontSizeCaseID = "@id-docx-direct-font-size-half-points"
const directRangeParsingCaseID = "@id-xlsx-go-direct-range-parsing"
const directRangeRefusalCaseID = "@id-xlsx-go-direct-range-refusal"
const formulaAnalysisCountsCaseID = "@id-xlsx-go-formula-analysis-counts"
const formulaQuotedSheetFlagsCaseID = "@id-xlsx-go-formula-quoted-sheet-flags"
const formulaAnalysisRefusalCaseID = "@id-xlsx-go-formula-analysis-refusal"
const formulaLiteralPunctuationCaseID = "@id-xlsx-go-formula-literal-punctuation"
const staticRemapExactCaseID = "@id-xlsx-go-static-remap-exact"
const staticRemapRefusalCaseID = "@id-xlsx-go-static-remap-refusal"
const staticReferencePropertiesCaseID = "@id-xlsx-go-static-reference-properties"
const retiredCacheCaseID = "@CACHE-001"

func selectedRunFormattingID(id string) bool {
	return id == runUnderlineCaseID || id == runFontNameCaseID || id == runColorCaseID || id == runHighlightCaseID || id == runVerticalAlignCaseID || id == runRoundtripFormattingCaseID
}

func selectedTableValueID(id string) bool {
	return id == tableDimensionsCaseID || id == tableCellAccessCaseID || id == tableCellTextCaseID || id == tableRowCountsCaseID
}

func xmlSelectedParsingCases() int {
	if xmlLexicalCandidate() || postBatch2Reference() {
		return 15
	}
	if xmlSafetyCandidate() {
		return 8
	}
	return 5
}

func selectedCanonicalID(path, id string) bool {
	if ((xmlLexicalCandidate() || postBatch2Reference()) && lexicalSelectedID(path, id)) || (contract20Candidate() && path == xmlEditingFeaturePath() && id == xmlApplyEditsID) {
		return true
	}
	return (path == overlapFeaturePath() && (id == bzipAdmissionCaseID || id == unsafeMembersCaseID)) || (path == packagePreservationFeaturePath() && (id == opcDetachedByteCaseID || id == opcCorpusNoopCaseID || (packageReasonsCandidate() && (id == opcReasonOpenID || id == opcReasonSaveID || id == opcReasonSymlinkID)) || (portableTransactionCandidate() && (id == opcDeferredTransactionID || id == opcOpaqueTransactionID)))) || (path == zip32FeaturePath() && packageReasonsCandidate() && (id == zipReasonReaderID || id == zipReasonWriterID || id == zipReasonBoundsID)) || (path == xmlParsingFeaturePath() && (id == xmlEntityValuesCaseID || id == xmlStylesheetPICaseID || id == xmlImplicitPrefixCaseID || id == xmlExpandedAttributeCaseID || id == xmlWhitespaceCaseID || (xmlSafetyCandidate() && (id == xmlPrototypeCaseID || id == xmlNamespaceCaseID || id == xmlMalformedCaseID)))) || (path == xmlEditingFeaturePath() && (id == attributeSpliceCaseID || id == attributeRefusalCaseID || id == elementRemovalCaseID || id == elementRemovalRefusalCaseID || id == elementReplacementCustodyCaseID || id == elementReplacementRefusalCaseID || id == childInsertionCustodyCaseID || id == childInsertionRefusalCaseID || id == childNamespaceMatrixCaseID)) || (path == xmlComparisonFeaturePath() && (id == xmlSignificantCaseID || id == xmlPrefixBindingCaseID || id == xmlUnsafeCaseID || id == xmlMarkupCaseID)) || id == canonicalID(path) || (path == formulaReferenceFeaturePath() && (id == directRangeRefusalCaseID || id == formulaAnalysisCountsCaseID || id == formulaQuotedSheetFlagsCaseID || id == formulaAnalysisRefusalCaseID || id == formulaLiteralPunctuationCaseID || id == staticRemapExactCaseID || id == staticRemapRefusalCaseID || id == staticReferencePropertiesCaseID)) || (path == runEffectsFeaturePath() && selectedRunFormattingID(id)) || (path == tableMergeFeaturePath() && (selectedTableValueID(id) || id == roundtripTableTextCaseID)) || (path == paragraphFeaturePath() && (id == paragraphAlignmentCaseID || id == paragraphSpacingCaseID || id == paragraphTogglesCaseID || id == paragraphRunsCaseID || id == bodyInsertOrderCaseID))
}

func canonicalID(path string) string {
	switch path {
	case xmlNamesFeaturePath():
		return unicodeQNameCaseID
	case xmlParsingFeaturePath():
		return xmlEntityValuesCaseID
	case xmlEditingFeaturePath():
		return immutableLeafCaseID
	case xmlComparisonFeaturePath():
		return xmlSignificantCaseID
	case overlapFeaturePath():
		return overlapCaseID
	case packagePreservationFeaturePath():
		return opcPreserveUnrelatedCaseID
	case zip32FeaturePath():
		if packageReasonsCandidate() {
			return zipReasonReaderID
		}
		return ""
	case negativeBudgetFeaturePath():
		return negativeBudgetCaseID
	case retainedVisibilityFeaturePath():
		return retainedVisibilityEvidenceCaseID
	case descriptorIntegrityFeaturePath():
		return descriptorCollisionCaseID
	case ownedChainFeaturePath():
		return ownedChainCaseID
	case crossSheetCacheFeaturePath():
		return crossSheetCacheCaseID
	case formulaReferenceFeaturePath():
		return directRangeParsingCaseID
	case runEffectsFeaturePath():
		return runEffectsCaseID
	case tableMergeFeaturePath():
		return tableMergeCaseID
	case creationFeaturePath():
		return newEmptyBodyCaseID
	case corePropertiesFeaturePath():
		return corePropertiesCaseID
	case pageLayoutFeaturePath():
		return sectionTitleBackgroundCaseID
	case paragraphFeaturePath():
		return paragraphTextGetterCaseID
	case directFontSizeFeaturePath():
		return directFontSizeCaseID
	default:
		return ""
	}
}

func inventoryFeature(path string, d os.DirEntry, err error, expected map[caseID]expectedCase, inventory *[]map[string]any, seen map[string]string, next func() string, canonical bool) error {
	if err != nil {
		return err
	}
	if (d != nil && d.IsDir()) || filepath.Ext(path) != ".feature" {
		return nil
	}
	f, err := os.Open(path)
	if err != nil {
		return err
	}
	doc, err := gherkin.ParseGherkinDocument(f, next)
	_ = f.Close()
	if err != nil {
		return err
	}
	if doc.Feature == nil {
		return fmt.Errorf("%s: no feature", path)
	}
	lifecycle, runner := "", ""
	for _, tag := range doc.Feature.Tags {
		switch tag.Name {
		case "@implemented", "@planned", "@external":
			if lifecycle != "" {
				return fmt.Errorf("%s: multiple lifecycles", path)
			}
			lifecycle = tag.Name
		case "@go", "@office":
			if runner != "" {
				return fmt.Errorf("%s: multiple runners", path)
			}
			runner = tag.Name
		}
	}
	if canonical {
		if canonicalID(path) == "" || lifecycle != "@planned" || runner != "" {
			return fmt.Errorf("%s: unexpected canonical feature metadata", path)
		}
	} else if lifecycle == "" || runner == "" {
		return fmt.Errorf("%s: lifecycle/runner missing", path)
	}
	if lifecycle == "@implemented" && runner != "@go" {
		return fmt.Errorf("%s: implemented case omitted by native runner", path)
	}
	if lifecycle == "@external" && runner != "@office" {
		return fmt.Errorf("%s: external runner mismatch", path)
	}
	lines := map[string]int{}
	ids := map[string]string{}
	uniformIDs := map[string]bool{}
	if canonical && uniformAPI18Candidate() && path == xmlEditingFeaturePath() {
		uniformIDs, err = uniformSelection()
		if err != nil {
			return err
		}
	}
	var collect func([]*messages.FeatureChild) error
	collect = func(children []*messages.FeatureChild) error {
		for _, c := range children {
			if c.Rule != nil {
				// Select the original custody Rule and the exact byte-custody
				// case in the second Rule; all its siblings remain unselected.
				if canonical && path == packagePreservationFeaturePath() && c.Rule.Name != "OPC package custody and transactional part edits" && c.Rule.Name != "OPC byte custody, transaction callbacks and safe save destinations" {
					continue
				}
				if !canonical || (path != retainedVisibilityFeaturePath() && path != packagePreservationFeaturePath() && path != zip32FeaturePath() && path != xmlEditingFeaturePath() && path != xmlParsingFeaturePath() && path != crossSheetCacheFeaturePath() && path != formulaReferenceFeaturePath() && path != runEffectsFeaturePath() && path != tableMergeFeaturePath() && path != creationFeaturePath() && path != corePropertiesFeaturePath() && path != pageLayoutFeaturePath() && path != paragraphFeaturePath() && path != directFontSizeFeaturePath()) {
					return fmt.Errorf("%s: Rules require explicit inventory support", path)
				}
				for _, child := range c.Rule.Children {
					if child.Background != nil {
						if path == packagePreservationFeaturePath() && c.Rule.Name == "OPC byte custody, transaction callbacks and safe save destinations" {
							continue // Exact selected Background is guarded below.
						}
						return fmt.Errorf("%s: rule backgrounds require explicit inventory support", path)
					}
					if child.Scenario != nil {
						if path == packagePreservationFeaturePath() && c.Rule.Name == "OPC byte custody, transaction callbacks and safe save destinations" && !hasScenarioTag(child.Scenario, opcDetachedByteCaseID) && !(packageReasonsCandidate() && (hasScenarioTag(child.Scenario, opcReasonOpenID) || hasScenarioTag(child.Scenario, opcReasonSaveID) || hasScenarioTag(child.Scenario, opcReasonSymlinkID) || (portableTransactionCandidate() && (hasScenarioTag(child.Scenario, opcDeferredTransactionID) || hasScenarioTag(child.Scenario, opcOpaqueTransactionID))))) {
							continue
						}
						if err := collect([]*messages.FeatureChild{{Scenario: child.Scenario}}); err != nil {
							return err
						}
					}
				}
				continue
			}
			if c.Background != nil {
				return fmt.Errorf("%s: backgrounds require explicit inventory support", path)
			}
			s := c.Scenario
			if s == nil {
				continue
			}
			if canonical && path == xmlEditingFeaturePath() && contract20Candidate() {
				if err := guardContract20ChildInsertionTags(s); err != nil {
					return err
				}
			}
			id := ""
			profile := ""
			for _, tag := range s.Tags {
				if nativeIDPattern.MatchString(tag.Name) || (canonical && selectedCanonicalID(path, tag.Name)) {
					if id != "" {
						return fmt.Errorf("%s: multiple IDs", path)
					}
					id = tag.Name
				} else if canonical && path == formulaReferenceFeaturePath() && tag.Name == "@profile-static-reference-api" && profile == "" {
					profile = tag.Name
				} else if canonical && path == overlapFeaturePath() && tag.Name == "@profile-physical-member-extents" && profile == "" {
					profile = tag.Name
				} else if canonical && path == packagePreservationFeaturePath() && (tag.Name == "@profile-opc-byte-custody" || packageReasonsCandidate() && tag.Name == "@profile-package-refusal-reasons" || portableTransactionCandidate() && tag.Name == "@profile-portable-transactions") && profile == "" {
					profile = tag.Name
				} else if canonical && path == zip32FeaturePath() && packageReasonsCandidate() && tag.Name == "@profile-zip32-refusal-reasons" && profile == "" {
					profile = tag.Name
				} else if canonical && path == xmlParsingFeaturePath() && (xmlLexicalCandidate() || postBatch2Reference()) && tag.Name == "@profile-xml-escaping-api" && profile == "" {
					profile = tag.Name
				} else if canonical && path == xmlEditingFeaturePath() && tag.Name == "@profile-lexical-snapshot-api" && profile == "" {
					profile = tag.Name
				} else if canonical && (path == creationFeaturePath() || path == corePropertiesFeaturePath() || path == pageLayoutFeaturePath() || path == paragraphFeaturePath()) && tag.Name == "@profile-document-value-api" && profile == "" {
					profile = tag.Name
				} else if canonical && path == tableMergeFeaturePath() && (tag.Name == "@profile-document-value-api" || tag.Name == "@profile-nullable-cell-api" || tag.Name == "@profile-bounded-cell-lookup" || tag.Name == "@profile-table-text-readback") && profile == "" {
					profile = tag.Name
				} else if canonical && path == runEffectsFeaturePath() && (tag.Name == "@profile-in-memory-effects-api" || tag.Name == "@profile-document-value-api" || tag.Name == "@profile-selected-formatting-readback") && profile == "" {
					profile = tag.Name
				} else if !canonical || selectedCanonicalID(path, id) {
					return fmt.Errorf("%s: unsupported scenario tag %s", path, tag.Name)
				}
			}
			if canonical && (!selectedCanonicalID(path, id) || uniformIDs[id]) {
				continue // Stronger uniform IDs are inventoried separately, not through historic guards.
			}
			if id == "" || (canonical && path == zip32FeaturePath() && packageReasonsCandidate() && profile != "@profile-zip32-refusal-reasons") || (canonical && path == packagePreservationFeaturePath() && ((id == opcDetachedByteCaseID && profile != "@profile-opc-byte-custody") || ((id == opcReasonOpenID || id == opcReasonSaveID || id == opcReasonSymlinkID) && profile != "@profile-package-refusal-reasons") || ((id == opcDeferredTransactionID || id == opcOpaqueTransactionID) && profile != "@profile-portable-transactions") || ((id == opcPreserveUnrelatedCaseID || id == opcCorpusNoopCaseID) && profile != ""))) || (canonical && path == xmlEditingFeaturePath() && profile != "@profile-lexical-snapshot-api" && id != xmlApplyEditsID) || (canonical && path == xmlParsingFeaturePath() && xmlLexicalCandidate() && id == xmlEscapingValuesID && profile != "@profile-xml-escaping-api") || (canonical && path == formulaReferenceFeaturePath() && profile != "@profile-static-reference-api") || (canonical && path == directFontSizeFeaturePath() && profile != "") || (canonical && path == overlapFeaturePath() && ((id == overlapCaseID && profile != "@profile-physical-member-extents") || ((id == bzipAdmissionCaseID || id == unsafeMembersCaseID) && profile != ""))) || (canonical && (path == creationFeaturePath() || path == corePropertiesFeaturePath() || path == pageLayoutFeaturePath() || path == paragraphFeaturePath()) && profile != "@profile-document-value-api") || (canonical && path == tableMergeFeaturePath() && ((id == tableCellAccessCaseID && profile != tableCellAccessProfile()) || (id == roundtripTableTextCaseID && profile != "@profile-table-text-readback") || (id != tableCellAccessCaseID && id != roundtripTableTextCaseID && profile != "@profile-document-value-api"))) || (canonical && path == runEffectsFeaturePath() && ((id == runEffectsCaseID && profile != "@profile-in-memory-effects-api") || ((id == runRoundtripFormattingCaseID && profile != "@profile-selected-formatting-readback") || (id != runEffectsCaseID && id != runRoundtripFormattingCaseID && profile != "@profile-document-value-api")))) {
				return fmt.Errorf("%s: scenario missing or mismatched ID/profile", path)
			}
			if !canonical && id == retiredChainCaseID {
				// Keep the historical source in the inventory, but select its canonical replacement.
				if path != filepath.Join(goFeatureRoot(), "implemented", "spreadsheet", "calc-chain.feature") || s.Name != "Dependent input edit atomically removes an owned calculation chain" || len(s.Steps) != 6 || len(s.Examples) != 0 {
					return fmt.Errorf("%s: retired CHAIN-001 source drift", path)
				}
			}
			if !canonical && id == retiredCacheCaseID {
				if path != filepath.Join(goFeatureRoot(), "implemented", "spreadsheet", "cache-invalidation.feature") || s.Name != "A cross-sheet chain invalidates transitively" || len(s.Steps) != 5 || len(s.Examples) != 0 {
					return fmt.Errorf("%s: retired CACHE-001 source drift", path)
				}
			}
			if prior, ok := seen[id]; ok {
				return fmt.Errorf("duplicate ID %s in %s and %s", id, prior, path)
			}
			seen[id] = path
			lines[s.Id] = int(s.Location.Line)
			ids[s.Id] = id
			if len(s.Steps) == 0 {
				return fmt.Errorf("%s: empty scenario", path)
			}
			for _, ex := range s.Examples {
				if len(ex.Tags) != 0 {
					return fmt.Errorf("%s: Examples tags not supported yet", path)
				}
				for _, row := range ex.TableBody {
					lines[row.Id] = int(row.Location.Line)
				}
			}
		}
		return nil
	}
	if err := collect(doc.Feature.Children); err != nil {
		return err
	}
	if canonical && path == xmlParsingFeaturePath() {
		if err := guardXMLEntityValuesRule(doc); err != nil {
			return err
		}
		if err := guardXMLWhitespaceRule(doc); err != nil {
			return err
		}
		if xmlSafetyCandidate() {
			if err := guardXMLSafetyRule(doc); err != nil {
				return err
			}
		}
	}
	if canonical && path == overlapFeaturePath() {
		if err := guardUnsafeMembersFeature(doc); err != nil {
			return err
		}
		if err := guardBZIPAdmissionRule(doc); err != nil {
			return err
		}
	}
	if canonical && (path == packagePreservationFeaturePath() || path == zip32FeaturePath()) && packageReasonsCandidate() {
		if err := guardReasonFeature(doc, path); err != nil {
			return err
		}
	}
	if canonical && path == packagePreservationFeaturePath() {
		if err := guardOPCPreserveUnrelatedRule(doc); err != nil {
			return err
		}
		if err := guardOPCCorpusNoopRule(doc); err != nil {
			return err
		}
		if err := guardOPCDetachedByteRule(doc); err != nil {
			return err
		}
	}
	if canonical && !(uniformAPI18Candidate() && path == xmlEditingFeaturePath()) && seen[canonicalID(path)] == "" {
		return fmt.Errorf("%s: selected canonical case missing", path)
	}
	if canonical && (xmlLexicalCandidate() || postBatch2Reference()) && (path == xmlParsingFeaturePath() || path == xmlNamesFeaturePath() || path == xmlEditingFeaturePath()) {
		if err := guardLexicalFeature(doc, path); err != nil {
			return err
		}
	}
	if canonical && path == xmlEditingFeaturePath() && !uniformAPI18Candidate() && seen[elementReplacementRefusalCaseID] == "" {
		return fmt.Errorf("%s: element-replacement refusal ID missing", path)
	}
	if canonical && path == xmlEditingFeaturePath() && !uniformAPI18Candidate() && seen[elementReplacementCustodyCaseID] == "" {
		return fmt.Errorf("%s: element-replacement custody ID missing", path)
	}
	if canonical && path == xmlEditingFeaturePath() && seen[childNamespaceMatrixCaseID] == "" {
		return fmt.Errorf("%s: child-namespace matrix ID missing", path)
	}
	if canonical && path == xmlComparisonFeaturePath() && (seen[xmlPrefixBindingCaseID] == "" || seen[xmlUnsafeCaseID] == "" || seen[xmlMarkupCaseID] == "") {
		return fmt.Errorf("%s: selected XML negative IDs missing", path)
	}
	if canonical && path == tableMergeFeaturePath() && boundedCellLookupCandidate() {
		if err := guardBoundedCellLookupRule(doc); err != nil {
			return err
		}
	}
	if canonical && path == tableMergeFeaturePath() && (seen[roundtripTableTextCaseID] == "" || seen[tableDimensionsCaseID] == "" || seen[tableCellAccessCaseID] == "" || seen[tableCellTextCaseID] == "" || seen[tableRowCountsCaseID] == "") {
		return fmt.Errorf("%s: selected table value cases missing", path)
	}
	if canonical && path == runEffectsFeaturePath() && (seen[runUnderlineCaseID] == "" || seen[runFontNameCaseID] == "" || seen[runColorCaseID] == "" || seen[runHighlightCaseID] == "" || seen[runVerticalAlignCaseID] == "" || seen[runRoundtripFormattingCaseID] == "") {
		return fmt.Errorf("%s: selected run-formatting outline missing", path)
	}
	canonicalCases := 0
	underlineRows := map[string]bool{}
	fontNameRows := map[string]bool{}
	colorRows := map[string]bool{}
	highlightRows := map[string]bool{}
	dimensionRows := map[string]bool{}
	directRangeParseRows := map[string]bool{}
	directRangeRefusalRows := map[string]bool{}
	formulaCountsRows := map[string]bool{}
	formulaRefusalRows := map[string]bool{}
	formulaFlagsRows := map[string]bool{}
	formulaLiteralRows := map[string]bool{}
	staticRemapExactRows := map[string]bool{}
	staticRemapRefusalRows := map[string]bool{}
	staticPropertiesRows := map[string]bool{}
	budgetRows := map[string]bool{}
	paragraphRows := map[string]bool{}
	alignmentRows := map[string]bool{}
	spacingRows := map[string]bool{}
	attributeSpliceRowsSeen := map[string]bool{}
	removalRefusalRowsSeen := map[string]bool{}
	replacementRefusalRowsSeen := map[string]bool{}
	for _, p := range gherkin.Pickles(*doc, path, next) {
		if canonical && (len(p.AstNodeIds) == 0 || !selectedCanonicalID(path, ids[p.AstNodeIds[0]])) {
			continue
		}
		line := lines[p.AstNodeIds[len(p.AstNodeIds)-1]]
		id := ids[p.AstNodeIds[0]]
		if canonical {
			canonicalCases++
			if path == packagePreservationFeaturePath() {
				var err error
				switch id {
				case opcPreserveUnrelatedCaseID:
					err = guardOPCPreserveUnrelatedCase(id, p, line)
				case opcCorpusNoopCaseID:
					err = guardOPCCorpusNoopCase(id, p, line)
				case opcDetachedByteCaseID:
					err = guardOPCDetachedByteCase(id, p, line)
				case opcReasonOpenID, opcReasonSaveID, opcReasonSymlinkID, opcDeferredTransactionID, opcOpaqueTransactionID:
					err = guardOPCReasonPickle(id, p, line)
				default:
					err = fmt.Errorf("unexpected selected OPC custody ID %s", id)
				}
				if err != nil {
					return err
				}
			}
			if path == zip32FeaturePath() && packageReasonsCandidate() {
				if err := guardZIPReasonPickle(id, p, line); err != nil {
					return err
				}
			}
			if path == overlapFeaturePath() && id == unsafeMembersCaseID {
				if err := guardUnsafeMembersCase(id, p, line); err != nil {
					return err
				}
			}
			if path == overlapFeaturePath() && id == bzipAdmissionCaseID {
				if err := guardBZIPAdmissionCase(id, p, line); err != nil {
					return err
				}
			}
			if path == xmlParsingFeaturePath() {
				var err error
				switch id {
				case xmlEntityValuesCaseID:
					err = guardXMLEntityValuesCase(id, p, line)
				case xmlStylesheetPICaseID:
					err = guardXMLStylesheetPICase(id, p, line)
				case xmlImplicitPrefixCaseID:
					err = guardXMLImplicitPrefixCase(id, p, line)
				case xmlExpandedAttributeCaseID:
					err = guardXMLExpandedAttributeCase(id, p, line)
				case xmlWhitespaceCaseID:
					err = guardXMLWhitespaceCase(id, p, line)
				case xmlPrototypeCaseID, xmlNamespaceCaseID, xmlMalformedCaseID:
					err = guardXMLSafetyCase(id, p, line)
				default:
					if xmlLexicalCandidate() || postBatch2Reference() {
						err = guardLexicalCase(id, p, line)
					} else {
						err = fmt.Errorf("unexpected selected XML parsing ID %s", id)
					}
				}
				if err != nil {
					return err
				}
			}
			if path == xmlNamesFeaturePath() {
				var err error
				if xmlLexicalCandidate() || postBatch2Reference() || (contract20Candidate() && id == unicodeQNameCaseID) {
					err = guardLexicalCase(id, p, line)
				} else {
					err = guardUnicodeQNameCase(id, p)
				}
				if err != nil {
					return err
				}
			}
			if path == xmlEditingFeaturePath() {
				if id == elementReplacementRefusalCaseID {
					row, err := guardElementReplacementRefusalCase(id, p, line)
					if err != nil {
						return err
					}
					if replacementRefusalRowsSeen[row] {
						return fmt.Errorf("%s: duplicate element-replacement refusal row %s", path, row)
					}
					replacementRefusalRowsSeen[row] = true
				} else if id == elementReplacementCustodyCaseID {
					if err := guardElementReplacementCustodyCase(id, p, line); err != nil {
						return err
					}
				} else if id == childNamespaceMatrixCaseID {
					if err := guardChildNamespaceMatrixCase(id, p, line); err != nil {
						return err
					}
				} else if id == childInsertionRefusalCaseID {
					if err := guardChildInsertionRefusalCase(id, p, line); err != nil {
						return err
					}
				} else if id == childInsertionCustodyCaseID {
					if err := guardChildInsertionCustodyCase(id, p, line); err != nil {
						return err
					}
				} else if id == elementRemovalCaseID {
					if err := guardElementRemovalCase(id, p, line); err != nil {
						return err
					}
				} else if id == elementRemovalRefusalCaseID {
					selection, err := guardElementRemovalRefusalCase(id, p, line)
					if err != nil {
						return err
					}
					if removalRefusalRowsSeen[selection] {
						return fmt.Errorf("%s: duplicate element-removal refusal row %s", path, selection)
					}
					removalRefusalRowsSeen[selection] = true
				} else if id == attributeRefusalCaseID {
					if err := guardAttributeRefusalCase(id, p, line); err != nil {
						return err
					}
				} else if id == attributeSpliceCaseID {
					row, err := guardAttributeSpliceCase(id, p, line)
					if err != nil {
						return err
					}
					if attributeSpliceRowsSeen[row] {
						return fmt.Errorf("%s: duplicate attribute-splice row %s", path, row)
					}
					attributeSpliceRowsSeen[row] = true
				} else if id == xmlApplyEditsID && (xmlLexicalCandidate() || postBatch2Reference() || contract20Candidate()) {
					if err := guardLexicalCase(id, p, line); err != nil {
						return err
					}
				} else if err := guardImmutableLeafCase(id, p); err != nil {
					return err
				}
			}
			if path == xmlComparisonFeaturePath() {
				if err := guardXMLComparisonCase(id, p); err != nil {
					return err
				}
			}
			if id == descriptorCollisionCaseID && (p.Name != "Unsigned descriptor geometry cannot excuse a corrupted payload" || len(p.Steps) != 8) {
				return fmt.Errorf("%s: unexpected descriptor-integrity case %q", path, p.Name)
			}
			if id == ownedChainCaseID && (p.Name != "Changing a precedent removes its owned nonstandard chain and invalidates dependent caches" || len(p.Steps) != 14) {
				return fmt.Errorf("%s: unexpected owned-chain case %q", path, p.Name)
			}
			if id == directRangeParsingCaseID || id == directRangeRefusalCaseID {
				if err := guardDirectRangeCase(id, p, path); err != nil {
					return err
				}
				rows := directRangeParseRows
				if id == directRangeRefusalCaseID {
					rows = directRangeRefusalRows
				}
				if rows[p.Name] {
					return fmt.Errorf("%s: duplicate direct-range row %q", path, p.Name)
				}
				rows[p.Name] = true
			}
			if id == formulaAnalysisCountsCaseID || id == formulaQuotedSheetFlagsCaseID || id == formulaAnalysisRefusalCaseID {
				if err := guardCanonicalFormulaAnalysisCase(id, p, path); err != nil {
					return err
				}
				rows := formulaCountsRows
				switch id {
				case formulaQuotedSheetFlagsCaseID:
					rows = formulaFlagsRows
				case formulaAnalysisRefusalCaseID:
					rows = formulaRefusalRows
				}
				if rows[p.Name] {
					return fmt.Errorf("%s: duplicate formula analysis row %q", path, p.Name)
				}
				rows[p.Name] = true
			}
			if id == formulaLiteralPunctuationCaseID || id == staticRemapExactCaseID || id == staticRemapRefusalCaseID || id == staticReferencePropertiesCaseID {
				if err := guardRemainingFormulaCase(id, p, path); err != nil {
					return err
				}
				rows := formulaLiteralRows
				switch id {
				case staticRemapExactCaseID:
					rows = staticRemapExactRows
				case staticRemapRefusalCaseID:
					rows = staticRemapRefusalRows
				case staticReferencePropertiesCaseID:
					rows = staticPropertiesRows
				}
				key := p.Name + "\x00" + p.Steps[0].Text
				if rows[key] {
					return fmt.Errorf("%s: duplicate remaining formula row %q", path, p.Name)
				}
				rows[key] = true
			}
			if id == crossSheetCacheCaseID && (p.Name != "An input edit invalidates a cached answer on another sheet" || len(p.Steps) != 14) {
				return fmt.Errorf("%s: unexpected cross-sheet cache case %q", path, p.Name)
			}
			if id == runEffectsCaseID && (p.Name != "Eight direct run effects read true after setting them" || len(p.Steps) != 3) {
				return fmt.Errorf("%s: unexpected run-effects case %q", path, p.Name)
			}
			if id == runUnderlineCaseID {
				style := strings.TrimSuffix(strings.TrimPrefix(p.Name, "A "), " underline is reflected by direct getters")
				if p.Name != "A "+style+" underline is reflected by direct getters" || !slices.Contains([]string{"single", "double", "thick", "dotted", "dash", "wave"}, style) || underlineRows[style] || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word run" || p.Steps[1].Text != "its underline style is set to "+style || p.Steps[2].Text != "Underline is true and UnderlineStyle equals "+style {
					return fmt.Errorf("%s: unexpected underline row %q", path, p.Name)
				}
				underlineRows[style] = true
			}
			if id == runFontNameCaseID {
				font := strings.TrimSuffix(strings.TrimPrefix(p.Name, "A run retains direct font name "), " in memory")
				if p.Name != "A run retains direct font name "+font+" in memory" || !slices.Contains([]string{"Arial", "Times New Roman", "Calibri", "Courier New", "Georgia", "Verdana"}, font) || fontNameRows[font] || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word run" || p.Steps[1].Text != "its font name is set to "+font || p.Steps[2].Text != "its font-name getter equals "+font {
					return fmt.Errorf("%s: unexpected font-name row %q", path, p.Name)
				}
				fontNameRows[font] = true
			}
			if id == runColorCaseID {
				variant := strings.TrimSuffix(strings.TrimPrefix(p.Name, "A "), " run colour is normalised by its getter")
				want := map[string][2]string{"red": {"FF0000", "FF0000"}, "hash-red": {"#FF0000", "FF0000"}, "lowercase": {"ff0000", "ff0000"}}
				values, ok := want[variant]
				if !ok || colorRows[variant] || p.Name != "A "+variant+" run colour is normalised by its getter" || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word run" || p.Steps[1].Text != "its colour is set to "+values[0] || p.Steps[2].Text != "its in-memory colour getter equals "+values[1] {
					return fmt.Errorf("%s: unexpected colour row %q", path, p.Name)
				}
				colorRows[variant] = true
			}
			if id == runHighlightCaseID {
				colour := strings.TrimSuffix(strings.TrimPrefix(p.Name, "A run retains highlight name "), " in memory")
				if p.Name != "A run retains highlight name "+colour+" in memory" || !slices.Contains([]string{"yellow", "cyan", "darkBlue", "lightGray", "black"}, colour) || highlightRows[colour] || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word run" || p.Steps[1].Text != "highlight is set to "+colour || p.Steps[2].Text != "the Highlight getter equals "+colour {
					return fmt.Errorf("%s: unexpected highlight row %q", path, p.Name)
				}
				highlightRows[colour] = true
			}
			if id == tableDimensionsCaseID {
				var rows, cols int
				if _, err := fmt.Sscanf(p.Name, "A newly added %d by %d table reports its dimensions", &rows, &cols); err != nil {
					return fmt.Errorf("%s: unexpected dimension row %q: %w", path, p.Name, err)
				}
				key := fmt.Sprintf("%dx%d", rows, cols)
				if !slices.Contains([]string{"1x1", "1x5", "5x1", "2x2", "3x3", "5x5", "10x3", "3x10"}, key) || dimensionRows[key] || p.Name != fmt.Sprintf("A newly added %d by %d table reports its dimensions", rows, cols) || len(p.Steps) != 3 || p.Steps[0].Text != "a new Word document" || p.Steps[1].Text != fmt.Sprintf("a table with %d rows and %d columns is added", rows, cols) || p.Steps[2].Text != fmt.Sprintf("RowCount equals %d and ColumnCount equals %d in memory", rows, cols) {
					return fmt.Errorf("%s: unexpected dimension row %q", path, p.Name)
				}
				dimensionRows[key] = true
			}
			if id == tableCellAccessCaseID || id == tableCellTextCaseID || id == tableRowCountsCaseID {
				var name string
				var steps []string
				switch id {
				case tableCellAccessCaseID:
					if boundedCellLookupCandidate() {
						if err := guardBoundedCellLookupCase(id, p, line); err != nil {
							return err
						}
						name = boundedCellLookupScenarioName
						steps = []string{
							"a new Word table has three rows and three columns with texts by row A1,B1,C1 then A2,B2,C2 then A3,B3,C3",
							"cells are looked up at these zero-based coordinates",
							"each lookup returns the listed presence and exact text without an exception",
							"the table still has three rows and three columns with its original texts and unchanged document XML",
						}
					} else {
						name = "A three-by-three table returns cells only at in-range coordinates"
						steps = []string{"a new Word table with three rows and three columns", "its Cell getter is called for all nine coordinates from zero through two", "each of those nine calls returns a nonnil cell", "calls for row or column negative one or three at the tested boundary coordinates return nil"}
					}
				case tableCellTextCaseID:
					name = "A two-by-two table reads four assigned texts and its first row"
					steps = []string{"a new Word table with two rows and two columns", "its cells are set by row to A1, B1, A2 and B2", "the four cell text getters equal A1, B1, A2 and B2 in those positions", "FirstRowText returns exactly A1 and B1"}
				case tableRowCountsCaseID:
					name = "Adding, inserting and deleting rows changes table count in memory"
					steps = []string{"a new Word table with two rows and three columns", "one row is appended, one is inserted at index one, and index one is deleted", "row counts after each step are three, four and three respectively", "deletion at index ten returns an error"}
				}
				if p.Name != name || len(p.Steps) != len(steps) {
					return fmt.Errorf("%s: unexpected table value case %q", path, p.Name)
				}
				for i, step := range steps {
					if p.Steps[i].Text != step {
						return fmt.Errorf("%s: unexpected table value step %d for %s", path, i+1, id)
					}
				}
			}
			if path == paragraphFeaturePath() {
				steps := make([]string, len(p.Steps))
				for i, step := range p.Steps {
					steps[i] = step.Text
				}
				var err error
				switch id {
				case paragraphTextGetterCaseID:
					err = guardParagraphTextRow(p.Name, steps, paragraphRows)
				case paragraphAlignmentCaseID:
					err = guardParagraphAlignmentRow(p.Name, steps, alignmentRows)
				case paragraphSpacingCaseID:
					err = guardParagraphSpacingRow(p.Name, steps, spacingRows)
				case paragraphTogglesCaseID, paragraphRunsCaseID:
					err = guardParagraphSingleCase(id, p.Name, steps)
				case bodyInsertOrderCaseID:
					err = guardBodyInsertOrderCase(p.Name, steps)
				}
				if err != nil {
					return fmt.Errorf("%s: %w", path, err)
				}
			}
			if id == directFontSizeCaseID {
				steps := make([]string, len(p.Steps))
				for i, step := range p.Steps {
					steps[i] = step.Text
				}
				if err := guardDirectFontSizeCase(p.Name, steps); err != nil {
					return fmt.Errorf("%s: %w", path, err)
				}
			}
			if id == corePropertiesCaseID || id == sectionTitleBackgroundCaseID {
				var name string
				var steps []string
				switch id {
				case corePropertiesCaseID:
					name = "Core properties read back three selected fields in memory"
					steps = []string{"a new Word document", "its core properties are set to title Doc Title, creator Doc Author and subject Doc Subject", "description Doc Description, keywords one;two, category Category and language en-US are supplied", "content status Draft, identifier urn:example:doc, last modifier Reviewer, revision 2 and version 1.0 are supplied", "created, modified and last-printed W3CDTF timestamps are supplied for 2026-02-03T00:00:00Z, 2026-02-03T01:00:00Z and 2026-02-03T02:00:00Z", "the setter and getter return no error", "only the in-memory title creator and subject are compared to Doc Title, Doc Author and Doc Subject"}
				case sectionTitleBackgroundCaseID:
					name = "First-section title page and document background read back in memory"
					steps = []string{"a new Word document with a first section", "TitlePage is set true on that section and BackgroundColor to EEEEEE", "the section TitlePage getter is true and the document BackgroundColor getter equals EEEEEE"}
				}
				if p.Name != name || len(p.Steps) != len(steps) {
					return fmt.Errorf("%s: unexpected property getter case %q", path, p.Name)
				}
				for i, step := range steps {
					if p.Steps[i].Text != step {
						return fmt.Errorf("%s: unexpected property getter step %d for %s", path, i+1, id)
					}
				}
			}
			if id == newEmptyBodyCaseID || id == roundtripTableTextCaseID {
				var name string
				var steps []string
				switch id {
				case newEmptyBodyCaseID:
					name = "A new document has a body and no paragraphs or tables"
					steps = []string{"a new Word document", "its body paragraphs and tables are enumerated", "the body is present with zero paragraphs and zero tables"}
				case roundtripTableTextCaseID:
					name = "Nine table cell texts survive save and reopen"
					steps = []string{"a new Word table with three rows and three columns", "its cells contain Header1, Header2, Header3, A1, B1, C1, A2, B2 and C2 in row order", "the document is saved and reopened", "exactly one table is readable", "all nine cell text getters equal their original row-order values"}
				}
				if p.Name != name || len(p.Steps) != len(steps) {
					return fmt.Errorf("%s: unexpected creation/readback case %q", path, p.Name)
				}
				for i, step := range steps {
					if p.Steps[i].Text != step {
						return fmt.Errorf("%s: unexpected creation/readback step %d for %s", path, i+1, id)
					}
				}
			}
			if id == tableMergeCaseID {
				steps := []string{"a new Word table with three rows and four columns", "cell zero-zero gets GridSpan three and first-column rows one and two get restart and continue", "GridSpan at zero-zero equals three", "VerticalMerge at row one is restart and at row two is continue"}
				if p.Name != "A cell span and two vertical-merge flags read back directly" || len(p.Steps) != len(steps) {
					return fmt.Errorf("%s: unexpected table merge getter case %q", path, p.Name)
				}
				for i, step := range steps {
					if p.Steps[i].Text != step {
						return fmt.Errorf("%s: unexpected table merge getter step %d", path, i+1)
					}
				}
			}
			if id == runVerticalAlignCaseID {
				steps := []string{"two new Word runs", "Superscript is enabled on the first and Subscript on the second", "the first reports superscript true and subscript false", "the second reports subscript true and superscript false"}
				if p.Name != "Separate superscript and subscript runs have opposite flags" || len(p.Steps) != len(steps) {
					return fmt.Errorf("%s: unexpected vertical-align case %q", path, p.Name)
				}
				for i, step := range steps {
					if p.Steps[i].Text != step {
						return fmt.Errorf("%s: unexpected vertical-align step %d", path, i+1)
					}
				}
			}
			if id == runRoundtripFormattingCaseID {
				steps := []string{"a new Word paragraph with three runs Bold-space, Italic-space and Colored", "the first run is bold, the second italic, and the third has colour FF0000, font size 14 and font Arial", "the document is saved and reopened", "at least one paragraph and three runs are readable", "the first run is bold and the second italic", "the third run reports colour FF0000, font size 14 and font Arial"}
				if p.Name != "Selected direct run formatting survives save and reopen" || len(p.Steps) != len(steps) {
					return fmt.Errorf("%s: unexpected selected formatting readback %q", path, p.Name)
				}
				for i, step := range steps {
					if p.Steps[i].Text != step {
						return fmt.Errorf("%s: unexpected selected formatting step %d", path, i+1)
					}
				}
			}
			if id == negativeBudgetCaseID {
				if (p.Name != "A negative source bytes budget refuses before package intake" && p.Name != "A negative entry count budget refuses before package intake") || budgetRows[p.Name] || len(p.Steps) != 6 {
					return fmt.Errorf("%s: unexpected negative-budget case %q", path, p.Name)
				}
				budgetRows[p.Name] = true
			}
			if id == retainedVisibilityEvidenceCaseID {
				if p.Name != "Retained four-slide inputs distinguish a namespaced marker from a visible control" || len(p.Steps) != 9 {
					return fmt.Errorf("%s: unexpected retained visibility evidence case %q", path, p.Name)
				}
			}
		}
		key := caseID{filepath.ToSlash(filepath.Clean(path)), id, line}
		if canonical {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": "@implemented", "runner": "@go", "selection": "canonical"})
		} else if id == retiredChainCaseID {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner, "selection": "superseded by " + ownedChainCaseID})
		} else if id == retiredCacheCaseID {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner, "selection": "superseded by " + crossSheetCacheCaseID})
		} else {
			*inventory = append(*inventory, map[string]any{"case": key, "name": p.Name, "lifecycle": lifecycle, "runner": runner})
		}
		if (lifecycle == "@implemented" && id != retiredChainCaseID && id != retiredCacheCaseID) || canonical {
			expected[key] = expectedCase{key, p.Name, len(p.Steps)}
		}
	}
	if canonical && ((canonicalID(path) == immutableLeafCaseID && !(uniformAPI18Candidate() && path == xmlEditingFeaturePath()) && (canonicalCases != xmlSelectedEditingCases() || len(attributeSpliceRowsSeen) != 3 || len(removalRefusalRowsSeen) != 2 || len(replacementRefusalRowsSeen) != 3)) || (canonicalID(path) == unicodeQNameCaseID && canonicalCases != xmlSelectedNamesCases()) || (canonicalID(path) == xmlEntityValuesCaseID && canonicalCases != xmlSelectedParsingCases()) || (canonicalID(path) == xmlSignificantCaseID && canonicalCases != 9) || (canonicalID(path) == negativeBudgetCaseID && (canonicalCases != 2 || len(budgetRows) != 2)) || (canonicalID(path) == retainedVisibilityEvidenceCaseID && canonicalCases != 1) || (canonicalID(path) == overlapCaseID && canonicalCases != 7) || (canonicalID(path) == opcPreserveUnrelatedCaseID && canonicalCases != packageSelectedReasonCases()) || (path == zip32FeaturePath() && packageReasonsCandidate() && canonicalCases != zipSelectedReasonCases()) || (canonicalID(path) == descriptorCollisionCaseID && canonicalCases != 1) || (canonicalID(path) == ownedChainCaseID && canonicalCases != 1) || (canonicalID(path) == crossSheetCacheCaseID && canonicalCases != 1) || (canonicalID(path) == directRangeParsingCaseID && (canonicalCases != 45 || len(directRangeParseRows) != 6 || len(directRangeRefusalRows) != 7 || len(formulaCountsRows) != 6 || len(formulaFlagsRows) != 1 || len(formulaRefusalRows) != 7 || len(formulaLiteralRows) != 5 || len(staticRemapExactRows) != 5 || len(staticRemapRefusalRows) != 7 || len(staticPropertiesRows) != 1)) || (canonicalID(path) == directFontSizeCaseID && canonicalCases != 1) || (canonicalID(path) == paragraphTextGetterCaseID && (canonicalCases != 16 || len(paragraphRows) != 5 || len(alignmentRows) != 4 || len(spacingRows) != 4)) || (canonicalID(path) == corePropertiesCaseID && canonicalCases != 1) || (canonicalID(path) == sectionTitleBackgroundCaseID && canonicalCases != 1) || (canonicalID(path) == newEmptyBodyCaseID && canonicalCases != 1) || (canonicalID(path) == tableMergeCaseID && (canonicalCases != 13 || len(dimensionRows) != 8)) || (canonicalID(path) == runEffectsCaseID && (canonicalCases != 23 || len(underlineRows) != 6 || len(fontNameRows) != 6 || len(colorRows) != 3 || len(highlightRows) != 5))) {
		return fmt.Errorf("%s: selected canonical case count drift: %d", path, canonicalCases)
	}
	return nil
}

func TestResultReconciliation(t *testing.T) {
	key := caseID{"f.feature", "@TEST-001", 12}
	expected := map[caseID]expectedCase{key: {key, "example", 1}}
	fixture := func(status string) []byte {
		return []byte(fmt.Sprintf(`[{"uri":"f.feature","elements":[{"name":"example","line":12,"type":"scenario","tags":[{"name":"@TEST-001"}],"steps":[{"name":"observed","result":{"status":%q}}]}]}]`, status))
	}
	cases := []struct {
		name      string
		data      []byte
		wantError bool
	}{
		{"passed", fixture("passed"), false}, {"missing", []byte(`[]`), true}, {"skipped", fixture("skipped"), true}, {"undefined", fixture("undefined"), true}, {"pending", fixture("pending"), true}, {"failed", fixture("failed"), true}, {"invalid", []byte(`no`), true},
	}
	var duplicated []reportFeature
	_ = json.Unmarshal(fixture("passed"), &duplicated)
	duplicated[0].Elements = append(duplicated[0].Elements, duplicated[0].Elements[0])
	dup, _ := json.Marshal(duplicated)
	cases = append(cases, struct {
		name      string
		data      []byte
		wantError bool
	}{"duplicate", dup, true})
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			err := reconcile(expected, tc.data)
			if (err != nil) != tc.wantError {
				t.Fatalf("error=%v expected failure=%v", err, tc.wantError)
			}
		})
	}
}
