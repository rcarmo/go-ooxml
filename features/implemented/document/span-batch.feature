@implemented @go @document
Feature: Atomic selected Word span batches
  A batch commits every selected replacement together or leaves all targets unchanged.

  @BATCH-001
  Scenario: Several matches in one leaf are replaced in one transaction
    Given selected Word spans for three occurrences of "red"
    When I replace all selected spans with "blue"
    Then the paragraph reads "blue blue blue"
    And every changed selected target is consumed

  @BATCH-002
  Scenario: An unsupported later target rolls back earlier selected edits
    Given selected Word spans with a later hyperlink-owned occurrence
    When I attempt a selected batch replacement with "blue"
    Then the selected batch refuses without changing the source archive
    And the first ordinary target remains reusable

  @BATCH-003
  Scenario: A stale selected target aborts the whole batch
    Given selected Word spans for three occurrences of "red"
    And the first occurrence has already been changed
    When I attempt a selected batch replacement with "blue"
    Then the stale batch leaves the previously committed edit unchanged
