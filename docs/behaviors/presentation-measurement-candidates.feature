@catalogue @not-executed
Feature: Native presentation measurement-entrypoint predicates
  Four benchmarks and one opt-in smoke test supply conditional error checks only.
  Reviewing these declarations does not execute a benchmark or measure memory or performance.

  @candidate-go-presentation-measurement-001
  Scenario: Minimal fixture read benchmark requires successful open and close when selected
    Given the benchmark runner selects BenchmarkPresentationOpenReader and the minimal fixture resolves
    When the fixture bytes are read and b.N iterations execute after ResetTimer
    Then each OpenReader and Close must succeed
    And a file-read error skips the benchmark while SetBytes records input size without a threshold assertion
    And ordinary go test without bench selection executes none of this benchmark body

  @candidate-go-presentation-measurement-002
  Scenario: Images fixture read benchmark has the same conditional open and close checks
    Given the benchmark runner selects BenchmarkPresentationOpenReaderLarge and the images fixture resolves
    When the fixture bytes are read and b.N iterations execute after ResetTimer
    Then each OpenReader and Close must succeed
    And a file-read error skips the benchmark without any assertion about fixture size or performance
    And neither decoded image content nor package equality is checked

  @candidate-go-presentation-measurement-003
  Scenario: Single-slide save benchmark checks only construction and each save result
    Given the selected benchmark constructs one slide with Benchmark text and a temporary output directory before ResetTimer
    When the same presentation is saved to a distinct bench-index path in each b.N iteration
    Then New and every SaveAs call must succeed
    And deferred Close errors are ignored and no saved file is reopened or compared
    And no time allocation or output-size threshold is asserted

  @candidate-go-presentation-measurement-004
  Scenario: Twenty-slide save benchmark leaves table and text content unverified
    Given the selected large-save benchmark constructs twenty slides each with a text box and two-by-two table whose first cell is Header
    When the same presentation is saved to distinct bench-large-index paths over b.N iterations
    Then New and every SaveAs call must succeed
    And deferred Close errors are ignored while no slide count text table or saved payload is read back
    And twenty setup iterations and b.N work iterations are not independent semantic test cases

  @candidate-go-presentation-measurement-005
  Scenario: Memory-profile entrypoint skips when disabled and otherwise checks fixture open and close
    Given the normal test runner discovers TestPresentationOpenReader_MemProfile
    When ENABLE_MEMPROFILE is empty
    Then the test skips before resolving or reading a fixture
    But when the variable is nonempty and the minimal fixture resolves the read OpenReader and Close must succeed
    And no profile is created by this body and no memory or allocation quantity is asserted
