@catalogue @not-executed
Feature: Native presentation fuzz seed predicates
  Bounded input seeds and conditional smoke paths are separate from schema and rendered-output claims.
  Existing mutable authoring calls do not establish the new canonical safe text-box workflow.

  @candidate-go-presentation-fuzz-001
  Scenario: Small generated decks retain their slide count after saving
    Given one to five requested slides and text at most 1024 bytes
    When a new presentation adds layout-zero slides with text boxes at fixed coordinates and saves and reopens
    Then reopening succeeds and SlideCount equals the requested count
    And the two seed tuples do not compare text box identity geometry text or formatting after reopening

  @candidate-go-presentation-fuzz-002
  Scenario: In-memory table construction retains requested dimensions
    Given one to eight rows and columns and text at most 512 bytes
    When a layout-zero slide receives a fixed-geometry table and the first cell is assigned text
    Then the table is nonnil and RowCount and ColumnCount equal the requests
    And the two seed tuples supply no cell-text readback save or reopen assertion

  @candidate-go-presentation-fuzz-003
  Scenario: Fixture mutation smoke paths permit operation errors without asserting equivalence
    Given eleven presentation fixture labels registered as seeds only after successful file reads with offset zero and XOR zero
    When nonempty input at most 8 MiB passes through MutateBytes and OpenReader succeeds
    Then Slides and Layouts accessors are invoked and saving is attempted
    And after successful save and reopen those accessors are invoked again
    And read errors omit seeds while open save or reopen errors return without failure assertions
    And no byte count content identity or typed-refusal check is made and zero-XOR seeds mutate no bytes
