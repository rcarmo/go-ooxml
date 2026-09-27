@catalogue @not-executed
Feature: Native parameterised spreadsheet predicates
  Twelve declarations dispatch finite local and helper-owned tables through dynamic subtest labels.
  Eighty selected table-row uses and two excluded rows describe source structure only.

  @candidate-go-spreadsheet-parameters-001
  Scenario: Nine generated workbook setups preserve their sheet count across save and reopen
    Given the empty multiple-sheets single-cell data-types formulas merged-cells large-data hidden-sheet and table-basic setups
    When New succeeds and each setup is saved and reopened with required successful SaveAs and Open results
    Then the reopened SheetCount equals the count captured after setup and before saving
    And cell contents formulas merges hidden flags and table data receive no readback assertion
    And setup mutation results and original and reopened Close errors are unchecked

  @candidate-go-spreadsheet-parameters-002
  Scenario: Seven value rows read back the expected in-memory cell type
    Given string int int64 float64 true false and nil values with their expected CellType enums
    When each value is assigned to A1 in a new workbook
    Then Type equals the expected enum
    And New and SetValue errors are ignored and value equality and persistence are not checked

  @candidate-go-spreadsheet-parameters-003
  Scenario: Seven selected cell-reference rows return the expected coordinates
    Given nine CommonCellRefCases rows with invalid and empty marked WantErr
    When WantErr rows are omitted before subtest dispatch and each remaining reference is passed to Cell
    Then the cell is nonnil and Row and Column equal the row expectations
    And the two omitted rows produce no invalid-input or refusal assertion

  @candidate-go-spreadsheet-parameters-004
  Scenario: Five selected range rows return the expected dimensions
    Given the single-cell row column block and large CommonRangeCases rows
    When each reference is passed to Range after excluding any WantErr rows
    Then the range is nonnil and RowCount and ColumnCount equal the inclusive endpoint differences
    And the current table contains no excluded rows and start coordinates and invalid-range behaviour are not asserted

  @candidate-go-spreadsheet-parameters-005
  Scenario: Six helper numeric values read back with exact float equality
    Given the zero positive-integer negative-integer float large and small CommonNumericCases rows
    When each float64 input is assigned to A1 and Float64 is called
    Then Float64 returns no error and its result equals Want exactly
    And SetValue errors cell type formatting and serialized output are unchecked

  @candidate-go-spreadsheet-parameters-006
  Scenario: Eight helper string values read back exactly in memory
    Given the empty simple spaced Unicode emoji special-character newline and tab CommonStringCases rows
    When each input is assigned to A1 and String is called
    Then String equals Want including the empty-string row
    And WantOK and SetValue errors are ignored and no save or cell-type check occurs

  @candidate-go-spreadsheet-parameters-007
  Scenario: Eight helper string values read back exactly after saving and reopening
    Given the eight CommonStringCases rows assigned to A1 in separate new workbooks
    When SaveAs succeeds and each path is reopened
    Then the reopened first sheet cell A1 String equals Want including the empty-string row
    And New Open SetValue and Close errors and WantOK are not asserted
    And cell type shared-string representation and unrelated package content are not compared

  @candidate-go-spreadsheet-parameters-008
  Scenario: Four sheet operations are checked through resulting count only
    Given rows for adding one adding two deleting Sheet2 by name and deleting index one
    When each operation runs after its requested initial sheet setup
    Then SheetCount equals two three two and two respectively
    And deletion errors sheet names order deleted identity and persistence are unchecked

  @candidate-go-spreadsheet-parameters-009
  Scenario: Twelve formula strings read back without calculation
    Given the four arithmetic SUM AVERAGE COUNT MAX MIN IF nested and cross-sheet formula rows
    When each formula is assigned to C1 in a new workbook
    Then Formula equals the supplied string and HasFormula is true and Type equals CellTypeFormula
    And SetFormula errors evaluation parsing validity caches and persistence are unchecked
    And the cross-sheet row does not create the referenced Sheet2

  @candidate-go-spreadsheet-parameters-010
  Scenario: Five row heights read back in memory
    Given heights ten fifteen twenty thirty and fifty
    When row one receives each height through SetHeight
    Then Height equals the supplied value exactly
    And no saved output bounds refusal or rendering is checked

  @candidate-go-spreadsheet-parameters-011
  Scenario: Five merge references read back from the in-memory collection
    Given row column block single-row-block and large merge references
    When each new first sheet receives one MergeCells call
    Then MergedCells has length one and its first Reference equals the supplied string
    And MergeCells errors overlap handling cell contents and persistence are unchecked
    And the first-element access is unconditional after the nonfatal length assertion

  @candidate-go-spreadsheet-parameters-012
  Scenario: Four UTC date inputs retain their calendar components in memory
    Given midnight dates 1900-01-01 2000-01-01 2024-06-15 and 2030-12-31
    When each date is assigned to A1 and Time is called
    Then Time returns no error and Year Month and Day equal the input components
    And SetValue errors time-of-day timezone serial representation date-system boundaries and persistence are unchecked
