package acceptance

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/wml"
)

var tableDimensions = [][2]int{{1, 1}, {1, 5}, {5, 1}, {2, 2}, {3, 3}, {5, 5}, {10, 3}, {3, 10}}
var cellTexts = [2][2]string{{"A1", "B1"}, {"A2", "B2"}}
var boundaryCells = [][2]int{{-1, 0}, {0, -1}, {3, 0}, {0, 3}, {3, 3}}
var boundedCellTexts = [3][3]string{{"A1", "B1", "C1"}, {"A2", "B2", "C2"}, {"A3", "B3", "C3"}}

func tableValueSteps(sc *godog.ScenarioContext, tableReadback *tableTextReadbackState) {
	var doc document.Document
	var table document.Table
	var rowCounts []int
	var observedCells [3][3]document.Cell
	var boundedObservations []boundedCellObservation
	var originalDocumentXML []byte
	var originalGrid []int64
	documentXML := func() ([]byte, error) {
		// The public Document interface does not expose XML; the native
		// implementation supplies this read of the same live model.
		xmlView, ok := doc.(interface{ XML() *wml.Document })
		if !ok {
			return nil, fmt.Errorf("Word document XML snapshot unavailable")
		}
		return xml.Marshal(xmlView.XML())
	}
	var deleteError error
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		doc, table, rowCounts, observedCells, boundedObservations, originalDocumentXML, originalGrid, deleteError = nil, nil, nil, [3][3]document.Cell{}, nil, nil, nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if doc != nil {
			_ = doc.Close()
		}
		return ctx, nil
	})
	newDocument := func() error {
		var err error
		doc, err = document.New()
		return err
	}
	newTable := func(rows, cols int) error {
		if doc == nil {
			if err := newDocument(); err != nil {
				return err
			}
		}
		table = doc.AddTable(rows, cols)
		if table == nil {
			return fmt.Errorf("new Word table is nil")
		}
		return nil
	}
	sc.Step(`^a new Word document$`, func() error {
		if err := newDocument(); err != nil {
			return err
		}
		tableReadback.doc = doc
		return nil
	})
	sc.Step(`^a table with (\d+) rows and (\d+) columns is added$`, func(rows, cols int) error {
		return newTable(rows, cols)
	})
	sc.Step(`^RowCount equals (\d+) and ColumnCount equals (\d+) in memory$`, func(rows, cols int) error {
		return checkTableDimensions(table, rows, cols)
	})
	sc.Step(`^a new Word table with (three|two) rows and (three|two) columns$`, func(rows, cols string) error {
		values := map[string]int{"two": 2, "three": 3}
		if err := newTable(values[rows], values[cols]); err != nil {
			return err
		}
		if rows == "three" && cols == "three" {
			tableReadback.doc, tableReadback.table = doc, table
		}
		return nil
	})
	sc.Step(`^its Cell getter is called for all nine coordinates from zero through two$`, func() error {
		if table == nil {
			return fmt.Errorf("Word table missing")
		}
		for row := range observedCells {
			for col := range observedCells[row] {
				observedCells[row][col] = table.Cell(row, col)
			}
		}
		return nil
	})
	sc.Step(`^each of those nine calls returns a nonnil cell$`, func() error {
		return checkObservedCells(observedCells)
	})
	sc.Step(`^calls for row or column negative one or three at the tested boundary coordinates return nil$`, func() error {
		return checkCellAccess(table, false)
	})
	sc.Step(`^a new Word table has three rows and three columns with texts by row A1,B1,C1 then A2,B2,C2 then A3,B3,C3$`, func() error {
		if err := newTable(3, 3); err != nil {
			return err
		}
		for row := range boundedCellTexts {
			for col, text := range boundedCellTexts[row] {
				cell := table.Cell(row, col)
				if cell == nil {
					return fmt.Errorf("cannot initialize cell (%d,%d)", row, col)
				}
				cell.SetText(text)
			}
		}
		grid := table.(interface{ XML() *wml.Tbl }).XML().TblGrid
		if grid == nil || len(grid.GridCol) != 3 {
			return fmt.Errorf("initial 3x3 table grid missing")
		}
		for _, col := range grid.GridCol {
			originalGrid = append(originalGrid, col.W)
		}
		var err error
		originalDocumentXML, err = documentXML()
		return err
	})
	sc.Step(`^cells are looked up at these zero-based coordinates$`, func(rows *godog.Table) error {
		if table == nil || originalDocumentXML == nil || !boundedCellLookupCandidate() {
			return fmt.Errorf("missing bounded table snapshot")
		}
		if err := guardBoundedCellLookupRuntimeTable(rows); err != nil {
			return err
		}
		boundedObservations = boundedObservations[:0]
		for i, entry := range boundedCellLookupRows {
			// Call the public API even for out-of-range positions. Do not clamp or
			// short-circuit here: an aliased boundary must be observed and fail.
			cell := table.Cell(entry.row, entry.col)
			observation := boundedCellObservation{row: entry.row, col: entry.col, present: cell != nil}
			if cell != nil {
				observation.text = cell.Text()
			}
			boundedObservations = append(boundedObservations, observation)
			if len(boundedObservations) != i+1 {
				return fmt.Errorf("lost lookup observation at row %d", i)
			}
		}
		return nil
	})
	sc.Step(`^each lookup returns the listed presence and exact text without an exception$`, func() error {
		return checkBoundedCellObservations(boundedObservations)
	})
	sc.Step(`^the table still has three rows and three columns with its original texts and unchanged document XML$`, func() error {
		if doc == nil || table == nil || len(boundedObservations) != len(boundedCellLookupRows) || len(originalDocumentXML) == 0 {
			return fmt.Errorf("bounded lookup did not complete")
		}
		if err := checkTableDimensions(table, 3, 3); err != nil {
			return err
		}
		grid := table.(interface{ XML() *wml.Tbl }).XML().TblGrid
		if grid == nil || len(grid.GridCol) != 3 || len(originalGrid) != 3 {
			return fmt.Errorf("table grid geometry changed")
		}
		for col, width := range originalGrid {
			if grid.GridCol[col].W != width {
				return fmt.Errorf("table grid column %d changed", col)
			}
		}
		for row := range boundedCellTexts {
			for col, text := range boundedCellTexts[row] {
				cell := table.Cell(row, col)
				if cell == nil || cell.Text() != text {
					return fmt.Errorf("table custody cell (%d,%d) changed", row, col)
				}
			}
		}
		current, err := documentXML()
		if err != nil || !bytes.Equal(current, originalDocumentXML) {
			return fmt.Errorf("document XML changed after cell lookup: %v", err)
		}
		return nil
	})
	sc.Step(`^its cells are set by row to A1, B1, A2 and B2$`, func() error {
		if table == nil {
			return fmt.Errorf("Word table missing")
		}
		for row := range cellTexts {
			for col, value := range cellTexts[row] {
				cell := table.Cell(row, col)
				if cell == nil {
					return fmt.Errorf("cell (%d,%d) missing", row, col)
				}
				cell.SetText(value)
			}
		}
		return nil
	})
	sc.Step(`^the four cell text getters equal A1, B1, A2 and B2 in those positions$`, func() error {
		return checkCellTexts(table)
	})
	sc.Step(`^FirstRowText returns exactly A1 and B1$`, func() error {
		return checkFirstRowText(table)
	})
	sc.Step(`^one row is appended, one is inserted at index one, and index one is deleted$`, func() error {
		if table == nil {
			return fmt.Errorf("Word table missing")
		}
		if table.AddRow() == nil {
			return fmt.Errorf("appended row is nil")
		}
		rowCounts = append(rowCounts, table.RowCount())
		if table.InsertRow(1) == nil {
			return fmt.Errorf("inserted row is nil")
		}
		rowCounts = append(rowCounts, table.RowCount())
		if err := table.DeleteRow(1); err != nil {
			return fmt.Errorf("delete row one: %w", err)
		}
		rowCounts = append(rowCounts, table.RowCount())
		deleteError = table.DeleteRow(10)
		return nil
	})
	sc.Step(`^row counts after each step are three, four and three respectively$`, func() error {
		return checkRowCounts(rowCounts)
	})
	sc.Step(`^deletion at index ten returns an error$`, func() error {
		return checkDeleteError(deleteError)
	})
}

func checkTableDimensions(table document.Table, rows, cols int) error {
	if table == nil {
		return fmt.Errorf("Word table missing")
	}
	if got := table.RowCount(); got != rows {
		return fmt.Errorf("RowCount = %d, want %d", got, rows)
	}
	if got := table.ColumnCount(); got != cols {
		return fmt.Errorf("ColumnCount = %d, want %d", got, cols)
	}
	return nil
}

func checkObservedCells(cells [3][3]document.Cell) error {
	for row := range cells {
		for col, cell := range cells[row] {
			if cell == nil {
				return fmt.Errorf("Cell(%d,%d) unexpectedly nil", row, col)
			}
		}
	}
	return nil
}

func checkCellAccess(table document.Table, inRange bool) error {
	if table == nil {
		return fmt.Errorf("Word table missing")
	}
	if inRange {
		var cells [3][3]document.Cell
		for row := range cells {
			for col := range cells[row] {
				cells[row][col] = table.Cell(row, col)
			}
		}
		return checkObservedCells(cells)
	}
	for _, coordinate := range boundaryCells {
		if table.Cell(coordinate[0], coordinate[1]) != nil {
			return fmt.Errorf("Cell(%d,%d) unexpectedly nonnil", coordinate[0], coordinate[1])
		}
	}
	return nil
}

func checkCellTexts(table document.Table) error {
	if table == nil {
		return fmt.Errorf("Word table missing")
	}
	for row := range cellTexts {
		for col, want := range cellTexts[row] {
			cell := table.Cell(row, col)
			if cell == nil {
				return fmt.Errorf("Cell(%d,%d) is nil", row, col)
			}
			if got := cell.Text(); got != want {
				return fmt.Errorf("Cell(%d,%d).Text = %q, want %q", row, col, got, want)
			}
		}
	}
	return nil
}

func checkFirstRowText(table document.Table) error {
	if table == nil {
		return fmt.Errorf("Word table missing")
	}
	got := table.FirstRowText()
	if len(got) != 2 || got[0] != "A1" || got[1] != "B1" {
		return fmt.Errorf("FirstRowText = %v, want [A1 B1]", got)
	}
	return nil
}

func checkRowCounts(got []int) error {
	want := []int{3, 4, 3}
	if len(got) != len(want) {
		return fmt.Errorf("row count observations = %v, want %v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			return fmt.Errorf("row count after operation %d = %d, want %d", i+1, got[i], want[i])
		}
	}
	return nil
}

func checkDeleteError(err error) error {
	if err == nil {
		return fmt.Errorf("DeleteRow(10) returned no error")
	}
	return nil
}

func TestTableValueNegativeControls(t *testing.T) {
	for _, dims := range tableDimensions {
		t.Run(fmt.Sprintf("dimensions/%dx%d", dims[0], dims[1]), func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			table := doc.AddTable(dims[0], dims[1])
			if err := checkTableDimensions(table, dims[0], dims[1]); err != nil {
				t.Fatal(err)
			}
			if err := checkTableDimensions(table, dims[0]+1, dims[1]); err == nil || !strings.Contains(err.Error(), "RowCount") {
				t.Fatalf("wrong row count passed: %v", err)
			}
			if err := checkTableDimensions(table, dims[0], dims[1]+1); err == nil || !strings.Contains(err.Error(), "ColumnCount") {
				t.Fatalf("wrong column count passed: %v", err)
			}
		})
	}
	t.Run("cell access", func(t *testing.T) {
		doc, err := document.New()
		if err != nil {
			t.Fatal(err)
		}
		defer doc.Close()
		table := doc.AddTable(3, 3)
		if err := checkCellAccess(table, true); err != nil {
			t.Fatal(err)
		}
		if err := checkCellAccess(table, false); err != nil {
			t.Fatal(err)
		}
		// Removing a row makes the selected nine-cell predicate fail.
		if err := table.DeleteRow(2); err != nil {
			t.Fatal(err)
		}
		if err := checkCellAccess(table, true); err == nil || !strings.Contains(err.Error(), "Cell") {
			t.Fatalf("missing in-range cell passed: %v", err)
		}
		// Growing to four rows makes the lower-boundary predicate fail.
		table.AddRow()
		table.AddRow()
		if err := checkCellAccess(table, false); err == nil || !strings.Contains(err.Error(), "Cell(3,0)") {
			t.Fatalf("available out-of-range cell passed: %v", err)
		}
	})
	t.Run("cell text and first row", func(t *testing.T) {
		doc, err := document.New()
		if err != nil {
			t.Fatal(err)
		}
		defer doc.Close()
		table := doc.AddTable(2, 2)
		for row := range cellTexts {
			for col, value := range cellTexts[row] {
				table.Cell(row, col).SetText(value)
			}
		}
		if err := checkCellTexts(table); err != nil {
			t.Fatal(err)
		}
		if err := checkFirstRowText(table); err != nil {
			t.Fatal(err)
		}
		for row := range cellTexts {
			for col, value := range cellTexts[row] {
				table.Cell(row, col).SetText("wrong")
				if err := checkCellTexts(table); err == nil || !strings.Contains(err.Error(), fmt.Sprintf("Cell(%d,%d)", row, col)) {
					t.Fatalf("wrong text %s at (%d,%d) passed: %v", value, row, col, err)
				}
				table.Cell(row, col).SetText(value)
			}
		}
		table.Cell(0, 0).SetText("wrong")
		if err := checkFirstRowText(table); err == nil || !strings.Contains(err.Error(), "FirstRowText") {
			t.Fatalf("wrong first row passed: %v", err)
		}
	})
	t.Run("row counts and deletion", func(t *testing.T) {
		if err := checkRowCounts([]int{3, 4, 3}); err != nil {
			t.Fatal(err)
		}
		for i := range []int{3, 4, 3} {
			wrong := []int{3, 4, 3}
			wrong[i]++
			if err := checkRowCounts(wrong); err == nil || !strings.Contains(err.Error(), "row count") {
				t.Fatalf("wrong operation count %d passed: %v", i, err)
			}
		}
		if err := checkDeleteError(fmt.Errorf("invalid index")); err != nil {
			t.Fatal(err)
		}
		if err := checkDeleteError(nil); err == nil || !strings.Contains(err.Error(), "DeleteRow(10)") {
			t.Fatalf("missing deletion error passed: %v", err)
		}
	})
}
