package acceptance

import (
	"bytes"
	"fmt"
	"os"
	"strconv"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
)

const boundedCellLookupScenarioName = "A three-by-three table reports exact cell presence without out-of-range aliasing"

type boundedCellLookupRow struct {
	row, col int
	present  bool
	text     string
}

type boundedCellObservation struct {
	row, col int
	present  bool
	text     string
}

var boundedCellLookupRows = []boundedCellLookupRow{
	{0, 0, true, "A1"}, {0, 1, true, "B1"}, {0, 2, true, "C1"},
	{1, 0, true, "A2"}, {1, 1, true, "B2"}, {1, 2, true, "C2"},
	{2, 0, true, "A3"}, {2, 1, true, "B3"}, {2, 2, true, "C3"},
	{-1, 0, false, ""}, {0, -1, false, ""}, {3, 0, false, ""},
	{0, 3, false, ""}, {3, 3, false, ""},
}

func tableCellAccessProfile() string {
	if boundedCellLookupCandidate() {
		return "@profile-bounded-cell-lookup"
	}
	return "@profile-nullable-cell-api"
}

func boundedCellLookupCandidate() bool {
	// Select by the sealed feature's authored contract, not a guessed pin
	// name. TestAcceptance verifies checkout HEAD, clean tree and manifest seal.
	data, err := os.ReadFile(tableMergeFeaturePath())
	return err == nil && containsBoundedCellMarker(data)
}

func containsBoundedCellMarker(data []byte) bool {
	return bytes.Contains(data, []byte("@profile-bounded-cell-lookup @id-docx-go-table-cell-access"))
}

func guardBoundedCellLookupTable(arg *messages.PickleStepArgument) error {
	if arg == nil || arg.DataTable == nil || len(arg.DataTable.Rows) != len(boundedCellLookupRows)+1 {
		return fmt.Errorf("bounded cell lookup table cardinality drift")
	}
	for i, row := range arg.DataTable.Rows {
		if len(row.Cells) != 4 {
			return fmt.Errorf("bounded cell lookup table width drift at %d", i)
		}
		want := [4]string{"row", "column", "present", "text"}
		if i > 0 {
			entry := boundedCellLookupRows[i-1]
			want = [4]string{strconv.Itoa(entry.row), strconv.Itoa(entry.col), strconv.FormatBool(entry.present), entry.text}
		}
		for col, cell := range row.Cells {
			if cell.Value != want[col] {
				return fmt.Errorf("bounded cell lookup table row %d col %d drift: %q != %q", i, col, cell.Value, want[col])
			}
		}
	}
	return nil
}

func guardBoundedCellLookupCase(id string, p *messages.Pickle, line int) error {
	steps := []string{
		"a new Word table has three rows and three columns with texts by row A1,B1,C1 then A2,B2,C2 then A3,B3,C3",
		"cells are looked up at these zero-based coordinates",
		"each lookup returns the listed presence and exact text without an exception",
		"the table still has three rows and three columns with its original texts and unchanged document XML",
	}
	if id != tableCellAccessCaseID || p == nil || line != 79 || p.Name != boundedCellLookupScenarioName || len(p.AstNodeIds) != 1 || len(p.Steps) != 4 || len(p.Tags) != 3 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != "@profile-bounded-cell-lookup" || p.Tags[2].Name != id {
		return fmt.Errorf("bounded cell lookup case identity drift: id=%s line=%d name=%q", id, line, p.Name)
	}
	for i, want := range steps {
		if p.Steps[i].Text != want || i != 1 && p.Steps[i].Argument != nil {
			return fmt.Errorf("bounded cell lookup step %d drift", i+1)
		}
	}
	return guardBoundedCellLookupTable(p.Steps[1].Argument)
}

func guardBoundedCellLookupRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "Word tables and cell properties" {
		return fmt.Errorf("bounded cell lookup feature drift")
	}
	seen := 0
	for _, child := range doc.Feature.Children {
		if child.Rule == nil {
			continue
		}
		for _, member := range child.Rule.Children {
			if member.Background != nil {
				continue
			}
			if member.Scenario == nil {
				continue
			}
			s := member.Scenario
			for _, tag := range s.Tags {
				if tag.Name != tableCellAccessCaseID {
					continue
				}
				seen++
				if len(s.Tags) != 2 || s.Tags[0].Name != "@profile-bounded-cell-lookup" || int(tag.Location.Line) != 78 || int(s.Location.Line) != 79 || len(s.Examples) != 0 || len(s.Steps) != 4 || s.Name != boundedCellLookupScenarioName || s.Steps[1].DataTable == nil || len(s.Steps[1].DataTable.Rows) != len(boundedCellLookupRows)+1 {
					return fmt.Errorf("bounded cell lookup authored structure drift")
				}
				if child.Rule.Name != "Document value API and selected save-reopen predicates" {
					return fmt.Errorf("bounded cell lookup rule drift: %q", child.Rule.Name)
				}
			}
		}
	}
	if seen != 1 {
		return fmt.Errorf("bounded cell lookup scenario count %d", seen)
	}
	return nil
}

func guardBoundedCellLookupRuntimeTable(table *godog.Table) error {
	if table == nil || len(table.Rows) != len(boundedCellLookupRows)+1 {
		return fmt.Errorf("bounded lookup runtime table cardinality drift")
	}
	for i, row := range table.Rows {
		if len(row.Cells) != 4 {
			return fmt.Errorf("bounded lookup runtime width drift at %d", i)
		}
		want := [4]string{"row", "column", "present", "text"}
		if i > 0 {
			e := boundedCellLookupRows[i-1]
			want = [4]string{strconv.Itoa(e.row), strconv.Itoa(e.col), strconv.FormatBool(e.present), e.text}
		}
		for j, cell := range row.Cells {
			if cell.Value != want[j] {
				return fmt.Errorf("bounded lookup runtime cell (%d,%d) drift", i, j)
			}
		}
	}
	return nil
}

func checkBoundedCellObservations(got []boundedCellObservation) error {
	if len(got) != len(boundedCellLookupRows) {
		return fmt.Errorf("bounded lookup observations %d, want %d", len(got), len(boundedCellLookupRows))
	}
	for i, want := range boundedCellLookupRows {
		observation := got[i]
		if observation.row != want.row || observation.col != want.col || observation.present != want.present || observation.text != want.text {
			return fmt.Errorf("Cell(%d,%d) present=%t text=%q, want %t/%q", want.row, want.col, observation.present, observation.text, want.present, want.text)
		}
	}
	return nil
}

func TestBoundedCellLookupGuardAndNegativeControls(t *testing.T) {
	loadReferencePin(t)
	if !boundedCellLookupCandidate() {
		t.Skip("published reference retains nullable-cell-api")
	}
	path := tableMergeFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil {
		t.Fatal(err)
	}
	if err := guardBoundedCellLookupRule(doc); err != nil {
		t.Fatal(err)
	}
	var selected *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == tableCellAccessCaseID {
				if selected != nil {
					t.Fatal("duplicate bounded cell case")
				}
				selected = p
			}
		}
	}
	if err := guardBoundedCellLookupCase(tableCellAccessCaseID, selected, 79); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name   string
		mutate func(*messages.Pickle)
	}{
		{"scenario name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"step", func(p *messages.Pickle) { s := *p.Steps[2]; s.Text += " changed"; p.Steps[2] = &s }},
		{"table entry", func(p *messages.Pickle) {
			s := *p.Steps[1]
			a := *s.Argument
			d := *a.DataTable
			d.Rows = append([]*messages.PickleTableRow(nil), d.Rows...)
			r := *d.Rows[1]
			r.Cells = append([]*messages.PickleTableCell(nil), r.Cells...)
			cell := *r.Cells[3]
			cell.Value = "wrong"
			r.Cells[3] = &cell
			d.Rows[1] = &r
			a.DataTable = &d
			s.Argument = &a
			p.Steps[1] = &s
		}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			clone := *selected
			clone.Steps = append([]*messages.PickleStep(nil), selected.Steps...)
			tc.mutate(&clone)
			if guardBoundedCellLookupCase(tableCellAccessCaseID, &clone, 79) == nil {
				t.Fatal("bounded cell lookup guard accepted drift")
			}
		})
	}
	for _, tc := range []struct {
		name   string
		mutate func([]boundedCellObservation)
	}{
		{"wrong in-range cell", func(g []boundedCellObservation) { g[5].text = "B1" }},
		{"negative aliased to cell", func(g []boundedCellObservation) { g[9].present, g[9].text = true, "A1" }},
		{"missing coordinate", func(g []boundedCellObservation) { g[0].text = "" }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			got := make([]boundedCellObservation, len(boundedCellLookupRows))
			for i, row := range boundedCellLookupRows {
				got[i] = boundedCellObservation{row.row, row.col, row.present, row.text}
			}
			tc.mutate(got)
			if checkBoundedCellObservations(got) == nil {
				t.Fatal("bounded cell lookup accepted wrong observation")
			}
		})
	}
}
