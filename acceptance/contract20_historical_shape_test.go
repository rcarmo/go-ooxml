package acceptance

import (
	"fmt"

	messages "github.com/cucumber/messages/go/v21"
)

// Historical shape is independent of the opt-in execution selections.
// Contract20 retains the API18 spelling and locations of existing cases;
// this must not enable batch-2 or the separately inventoried API18 lane.
func historicalAPI18Shape() bool {
	return postBatch2Reference() || contract20Candidate()
}

// The API18 tail adds lines at different positions in editing.feature.
// Keep each existing lexical/default line and pin the Contract20 line per case.
func historicalEditingLine(defaultLine, lexicalLine, contract20Line int) int {
	if contract20Candidate() {
		return contract20Line
	}
	return lexicalEditingLine(defaultLine, lexicalLine)
}

const uniformHistoricalAssertion = "the uniform profile result, refusal category and immutable input custody match the sealed API contract"

// This selected historical case has an additional authored-value step. Check
// its profile and stable ID before the canonical inventory can skip unknown IDs.
func guardContract20ChildInsertionTags(s *messages.Scenario) error {
	if s == nil || s.Name != "Insert a child with independently scoped element and attribute names" {
		return nil
	}
	if len(s.Tags) != 2 || s.Tags[0].Name != "@profile-lexical-snapshot-api" || s.Tags[1].Name != childInsertionCustodyCaseID {
		return fmt.Errorf("Contract20 child insertion profile/ID drift")
	}
	return nil
}

// Preserve each existing guard's prefix comparison, and require exactly one
// additional, argument-free historical assertion for the Contract20 checkout.
func historicalStepCount(p *messages.Pickle, oldCount int) bool {
	if !contract20Candidate() {
		return len(p.Steps) == oldCount
	}
	if len(p.Steps) != oldCount+1 || p.Steps[oldCount].Text != uniformHistoricalAssertion || p.Steps[oldCount].Argument != nil {
		return false
	}
	for _, step := range p.Steps[:oldCount] {
		if step.Argument != nil {
			return false
		}
	}
	return true
}

// Formula-analysis successes and direct-range parses have an independent,
// source-specific normalized JSON assertion before the uniform tail. Match
// the complete sealed vector by ID and expanded row name; retain the old
// guard's independently compiled prefix and refusal checks.
func historicalNormalizedFormulaSteps(id string, p *messages.Pickle, oldCount int) bool {
	if !contract20Candidate() {
		return len(p.Steps) == oldCount
	}
	if id != formulaAnalysisCountsCaseID && id != formulaQuotedSheetFlagsCaseID && id != formulaLiteralPunctuationCaseID && id != directRangeParsingCaseID {
		return historicalStepCount(p, oldCount)
	}
	l, err := readUniformLedger()
	if err != nil {
		return false
	}
	for _, file := range l.Files {
		if file.Path != "workflows/xlsx/formula-references.feature" {
			continue
		}
		for _, scenario := range file.Scenarios {
			if scenario.ID != id {
				continue
			}
			for _, row := range scenario.After {
				if row.Name != p.Name || len(row.Steps) != oldCount+2 || len(p.Steps) != len(row.Steps) {
					continue
				}
				for i, step := range row.Steps {
					if p.Steps[i].Text != step.Text || p.Steps[i].Argument != nil {
						return false
					}
				}
				return p.Steps[len(p.Steps)-1].Text == uniformHistoricalAssertion
			}
		}
	}
	return false
}
