package acceptance

import (
	"fmt"

	messages "github.com/cucumber/messages/go/v21"
)

// Keep the selected rows and complete steps tied to the canonical feature,
// independently of the production analyser and the step bindings.
func guardRemainingFormulaCase(id string, p *messages.Pickle, path string) error {
	literal := map[string]struct{ source, result string }{
		"adjacent string operator": {`A1 "+" B1`, "refused with zero references"},
		"adjacent string range":    {`A1 ":" B2`, "refused with zero references"},
		"quoted close in SUM":      {`SUM(")",A1)`, "accepted without error"},
		"quoted percent in SUM":    {`SUM("%",A1)`, "accepted without error"},
		"quoted equality only":     {`"="`, "accepted without error"},
	}
	type remapRow struct{ axis, at, count, sheet, source, expected string }
	exact := map[string]remapRow{
		"row at 2 by 1":    {"row", "2", "1", "Main", `IF(A1="A2",A2,Other!A2)`, `IF(A1="A2",A3,Other!A2)`},
		"column at 2 by 1": {"column", "2", "1", "Main", `'O''Brien'!$b$2 + Main!c1`, `'O''Brien'!$b$2 + Main!D1`},
		"row at 3 by 2":    {"row", "3", "2", "Main", `SUM(A5:A1)`, `SUM(A7:A1)`},
	}
	refusal := map[string]remapRow{
		"invalid axis":      {"bad", "1", "1", "Main", `A1+B1`, ""},
		"zero coordinate":   {"row", "0", "1", "Main", `A1+B1`, ""},
		"zero count":        {"column", "1", "0", "Main", `A1+B1`, ""},
		"blank sheet":       {"row", "1", "1", "", `A1+B1`, ""},
		"row overflow":      {"row", "1", "1", "Main", `A1048576`, ""},
		"column overflow":   {"column", "1", "1", "Main", `XFD1`, ""},
		"dynamic reference": {"row", "1", "1", "Main", `INDIRECT(A1)`, ""},
	}
	var steps []string
	switch id {
	case formulaLiteralPunctuationCaseID:
		for variant, row := range literal {
			if p.Name == variant+" string punctuation has a bounded parse outcome" {
				steps = []string{"the formula source is JSON " + canonicalFormulaJSONString(row.source), "the static formula analyser reads the source", "the analysis is " + row.result}
			}
		}
	case staticRemapExactCaseID:
		for _, row := range []remapRow{exact["row at 2 by 1"], exact["column at 2 by 1"], exact["row at 3 by 2"],
			{"row", "3", "2", "Main", `a1 + Other!b2`, `a1 + Other!b2`},
			{"row", "2", "1", "main", `Main!A1:A3`, `Main!A1:A4`}} {
			if p.Name == fmt.Sprintf("Inserting a %s at %s by %s rewrites only supported references", row.axis, row.at, row.count) && len(p.Steps) > 0 && p.Steps[0].Text == "the formula source is JSON "+canonicalFormulaJSONString(row.source)+" in context sheet Main" {
				steps = remapCanonicalSteps(row.axis, row.at, row.count, row.sheet, row.source, row.expected, true)
			}
		}
	case staticRemapRefusalCaseID:
		for variant, row := range refusal {
			if p.Name == "A "+variant+" insertion refuses without a partial replacement" {
				steps = remapCanonicalSteps(row.axis, row.at, row.count, row.sheet, row.source, "", false)
			}
		}
	case staticReferencePropertiesCaseID:
		if p.Name == "Static reference slices and unaffected remaps remain stable in a finite expression matrix" {
			steps = []string{"cell tokens A1, $B2, C$3 and $XFD$9", "optional prefixes empty, Main! and 'Input Data'! with operators +, -, *, /, & and >=", "the static analyser checks all 3 by 4 by 4 by 6 source expressions", "every expression has two references whose source slices each parse as one matching reference", "inserting one row at 100 on Main leaves each original expression byte-identical", "wrapping each expression in SUM preserves its references after adjusting their byte spans"}
		}
	}
	if len(steps) == 0 || !historicalNormalizedFormulaSteps(id, p, len(steps)) {
		return fmt.Errorf("%s: unexpected canonical formula row %s %q", path, id, p.Name)
	}
	for i, step := range steps {
		if p.Steps[i].Text != step {
			return fmt.Errorf("%s: canonical formula row %s %q step %d drift: got %q, want %q", path, id, p.Name, i+1, p.Steps[i].Text, step)
		}
	}
	return nil
}

func remapCanonicalSteps(axis, at, count, sheet, source, expected string, success bool) []string {
	steps := []string{"the formula source is JSON " + canonicalFormulaJSONString(source) + " in context sheet Main", "the static remapper inserts " + axis + " at " + at + " by " + count + " on sheet JSON " + canonicalFormulaJSONString(sheet)}
	if success {
		return append(steps, "the complete replacement expression equals JSON "+canonicalFormulaJSONString(expected))
	}
	return append(steps, "it returns an error and an empty replacement expression")
}
