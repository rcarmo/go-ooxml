@catalogue @not-executed
Feature: Native spreadsheet measurement-entrypoint predicates
  Four benchmarks and one opt-in smoke test supply conditional error checks only.
  Source review executes no benchmark and measures no memory or performance.

  @candidate-go-spreadsheet-measurement-001
  Scenario: Minimal workbook benchmark checks open and close when selected
    Given the runner selects BenchmarkWorkbookOpenReader and the minimal fixture resolves
    When the fixture is read and b.N iterations run after SetBytes and ResetTimer
    Then every OpenReader and Close must succeed
    And a file-read error skips the benchmark while resolver failures precede that branch
    And input size is recorded without a threshold and workbook contents are not asserted

  @candidate-go-spreadsheet-measurement-002
  Scenario: Tables workbook benchmark checks open and close without content comparison
    Given the runner selects BenchmarkWorkbookOpenReaderLarge and the tables fixture resolves
    When the fixture is read and b.N iterations run after SetBytes and ResetTimer
    Then every OpenReader and Close must succeed
    And a file-read error skips the benchmark without asserting source size or performance
    And no table content or package equality is checked

  @candidate-go-spreadsheet-measurement-003
  Scenario: Small save benchmark checks construction and saves while ignoring setup results
    Given the selected benchmark constructs one workbook and sets A1 to Benchmark and B2 to 123.45 before ResetTimer
    When the same workbook is saved to a distinct bench-index path in each b.N iteration
    Then New and every SaveAs must succeed
    And SetValue results and deferred Close errors are ignored
    And saved files are not reopened and no time allocation or output-size threshold is asserted

  @candidate-go-spreadsheet-measurement-004
  Scenario: Large save benchmark leaves populated cells and table contents unverified
    Given the selected benchmark attempts three cell values in each of five hundred rows and creates BenchmarkTable over A1:C501 with a first-row update before ResetTimer
    When the same workbook is saved to distinct bench-large-index paths over b.N iterations
    Then New and every SaveAs must succeed
    And all fifteen hundred SetValue results the UpdateRow result and deferred Close errors are ignored
    And no cell count table content saved payload or performance quantity is asserted
    And setup counts and b.N iterations add no independent semantic cases

  @candidate-go-spreadsheet-measurement-005
  Scenario: Memory-profile entrypoint skips when empty and checks fixture open and close when enabled
    Given the normal test runner discovers TestWorkbookOpenReader_MemProfile
    When ENABLE_MEMPROFILE is empty including an explicitly empty value
    Then the test skips before fixture resolution
    But when the variable is nonempty and the minimal fixture resolves the read OpenReader and Close must succeed
    And the body creates no profile and asserts no memory allocation or workbook content quantity
