@catalogue @not-executed
Feature: Native spreadsheet fuzz seed predicates
  These targets describe bounded seed paths and conditional checks, not an exploratory fuzz campaign.
  API/helper agreement and successful reopen do not prove semantic equality or retained-source custody.

  @candidate-go-spreadsheet-fuzz-001
  Scenario: A returned cell agrees with the reference parser
    Given a new workbook and a cell-reference input from the cell-access target
    When the first raw sheet returns a nonnil Cell
    Then ParseCellRef must succeed and its Row and Col equal the returned cell coordinates
    And a nil Cell returns without a failure assertion
    And four literal seeds supply no independent coordinate oracle or persistence check

  @candidate-go-spreadsheet-fuzz-002
  Scenario: A returned range agrees with parser dimensions
    Given a new workbook and a range-reference input from the range-access target
    When the first raw sheet returns a nonnil Range
    Then ParseRangeRef must succeed and RowCount and ColumnCount match its dimensions
    And a nil Range returns without a failure assertion
    And the three seeds do not compare exact endpoints cell values or serialized output

  @candidate-go-spreadsheet-fuzz-003
  Scenario: Workbook save and reopen retains at least one sheet
    Given a nonempty sheet name at most 64 bytes a cell reference at most 16 bytes and a value at most 256 bytes
    When a new workbook adds the sheet and optionally calls SetValue on a nonnil cell
    Then saving and reopening succeed and the reopened workbook has at least one sheet
    And SetValue errors are ignored and neither the requested sheet nor its value is read back
    And two seed tuples add no sheet-name or value-equivalence assertion

  @candidate-go-spreadsheet-fuzz-004
  Scenario: Accepted table construction leaves some table after reopen
    Given a nonempty table reference at most 32 bytes and a name at most 64 bytes
    When the first sheet returns a nonnil table from AddTable
    Then after an unchecked empty-row append saving and reopening succeed and Tables is nonempty
    And a nil table returns without a failure assertion
    And neither appended-row outcome table name geometry nor cell data is compared

  @candidate-go-spreadsheet-fuzz-005
  Scenario: Fixture mutation smoke paths tolerate operation errors without asserting equivalence
    Given eleven fixture labels whose bytes are added as seeds only after successful file reads with offset zero and XOR zero
    When nonempty input at most 8 MiB passes through MutateBytes and OpenReader succeeds
    Then Sheets and Tables accessors are invoked and saving is attempted
    And after successful save and reopen those accessors are invoked again
    And open save or reopen errors return without a failure assertion while read errors omit individual seeds
    And there is no value count byte-equality or typed-refusal assertion and zero-XOR seeds change no bytes
