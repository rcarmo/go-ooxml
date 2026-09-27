@catalogue @not-executed
Feature: Native spreadsheet style admission and image target predicates
  These two declarations test selected validation and in-memory edit outcomes.
  Synthetic fixtures and receipt checks do not establish rendered output or complete package preservation.

  @candidate-go-spreadsheet-targets-001
  Scenario: Implicit style zero accepts absent styles or an actual format entry
    Given one fixture without a styles part and another with one actual xf but no count attribute
    When both open as editing sessions and ValidateStyles is called
    Then both implicit-zero cells pass validation
    And no numeric edit or style rendering is performed

  @candidate-go-spreadsheet-targets-002
  Scenario: Six style-index or table ambiguities refuse validation
    Given rows for nonzero without styles empty table with explicit or implicit zero overflow index empty index and duplicate cellXfs tables
    When each fixture opens and ValidateStyles is called
    Then each row returns a nonnil validation error
    And no exact error kind row-column style coverage or mutation rollback is asserted

  @candidate-go-spreadsheet-targets-003
  Scenario: Image no-op and replacement retain distinct target lifecycle rules
    Given a two-by-two PNG picture and an existing case-colliding IMAGE1.jpg media name
    When the original PNG bytes are supplied to ReplaceImage
    Then its receipt has no changes and the target is unconsumed
    And another session refuses that target while its owning session accepts JPEG replacement
    And the resulting graph has xl/media/image2.jpg with JPEG content type
    And the consumed target refuses reuse while a fresh target accepts the same JPEG bytes
    And the test does not compare complete archive drawing geometry original media or fresh no-op receipt bytes

  @candidate-go-spreadsheet-targets-004
  Scenario: Truncated JPEG refusal leaves its image target reusable
    Given a fresh target for the synthetic PNG picture
    When half of the replacement JPEG payload is supplied
    Then replacement returns an error with target consumed false and no receipt changes
    And the full JPEG then succeeds through the same target
    And no exact refusal kind or independent package-byte comparison is asserted

  @candidate-go-spreadsheet-targets-005
  Scenario: One image replacement invalidates another held handle
    Given two targets independently found for the same picture in one session
    When replacement through the first target succeeds
    Then replacement through the second target returns an error
    And this subtest asserts no exact error kind or separate post-refusal byte comparison
