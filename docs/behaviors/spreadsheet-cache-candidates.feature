@catalogue @not-executed
Feature: Native static cache invalidation lifecycle assertions
  These staging outcomes describe three native declarations and their local helpers.
  Report counts and selected byte comparisons do not calculate formulas or certify complete cache state.

  @candidate-go-spreadsheet-cache-001
  Scenario: A same-value write preserves the worksheet and keeps its target reusable
    Given S A1 stores 1 and two formula cells depend transitively on its static range including an empty cache
    When SetNumberWithInvalidation writes 1 then writes 5 using the same target
    Then the first call reports no ValueChanged and the worksheet bytes remain identical
    And the second call reports two Invalidated entries and workbook bytes contain fullCalcOnLoad set to 1
    And this subtest does not independently inspect removed cache nodes or invalidated addresses

  @candidate-go-spreadsheet-cache-002
  Scenario: A cyclic static dependency closure terminates with two reported entries
    Given B1 references A1 and C1 while C1 references B1
    When SetNumberWithInvalidation changes A1 from 1 to 3
    Then the call succeeds and reports two Invalidated entries
    And no independent calculation or exact cache-node assertion is made

  @candidate-go-spreadsheet-cache-003
  Scenario: A string-literal formula does not add an invalidation entry
    Given B1 has a string formula whose literal text is A1
    When SetNumberWithInvalidation changes A1 from 1 to 3
    Then the call succeeds with no Invalidated entries and State caches-unchanged
    And the unchanged cache bytes are not independently compared in this subtest

  @candidate-go-spreadsheet-cache-004
  Scenario: Array and volatile formula refusal preserves selected worksheet and target state
    Given separate fixtures with an array formula or NOW plus an A1 reference
    When SetNumberWithInvalidation attempts to change A1 from 1 to 3
    Then each call returns an error while the worksheet bytes remain identical and target consumed is false
    And no exact error kind or successful subsequent write is asserted in these two rows

  @candidate-go-spreadsheet-cache-005
  Scenario: A quoted sheet and full-grid reference invalidate one reported entry
    Given a formula references the entire A1 to XFD1048576 range on sheet O'Brien
    When SetNumberWithInvalidation changes that sheet's A1 to 7
    Then the call succeeds and reports one Invalidated entry
    And a second call with the same target refuses with typed stale_target
    And the test does not measure allocation complexity or verify all grid cells

  @candidate-go-spreadsheet-cache-006
  Scenario: Nonfinite inputs refuse before a later valid use of the same target
    Given a cross-sheet static dependency on O'Brien A1
    When NaN positive infinity and negative infinity are attempted on the same target
    Then every call errors with selected input-worksheet bytes unchanged and target consumed false
    And a subsequent finite write of 4 using that target succeeds
    And no full-package byte comparison or exact nonfinite refusal kind is asserted

  @candidate-go-spreadsheet-cache-007
  Scenario: A target from another editing session refuses
    Given two independently opened sessions with the same cross-sheet dependency
    When one session receives a numeric target from the other
    Then SetNumberWithInvalidation returns an error
    And this subtest makes no exact refusal-kind or rollback byte assertion

  @candidate-go-spreadsheet-cache-008
  Scenario: A nonstandard calculation chain is removed once and a later edit succeeds
    Given a workbook-owned calculation chain at custom/order.xml and a dependent formula cell
    When SetNumberWithInvalidation changes A1 and a fresh A1 target is edited again
    Then the first effect names custom/order.xml and its receipt has a delete entry with empty AfterSHA256
    And the package graph validates and the second effect names no removed chain
    And no disk reopen or independent chain-cache XML readback is performed in this declaration

  @candidate-go-spreadsheet-cache-009
  Scenario: Mixed chain content and signatures refuse without consuming the target
    Given separate fixtures with mixed calculation-chain text or a signature part
    When SetNumberWithInvalidation attempts to change A1
    Then each call errors and before and after serialization buffers are identical with target consumed false
    And the test ignores serialization errors and does not assert a specific refusal kind

  @candidate-go-spreadsheet-cache-010
  Scenario: The fixture helper requires an unresolved chain relationship to refuse at intake
    Given the helper builds a workbook chain relationship to absent custom/missing.xml
    When the helper calls OpenEditing
    Then the helper requires a nonnil error and returns nil
    And the caller adds no separate error-kind target-lifecycle or byte-preservation assertion
