@implemented @go @spreadsheet
Feature: Static dependency cache invalidation without calculation
  A supported input edit removes affected cached results and preserves unrelated caches.

  @CACHE-001
  Scenario: A cross-sheet chain invalidates transitively
    Given a workbook with input one and a static cross-sheet formula chain
    When I set the input to ten with explicit invalidation
    Then direct and transitive cached results are absent or empty
    And unrelated cached results and all formula text remain unchanged
    And the workbook requests full recalculation without computing an answer

  @CACHE-002
  Scenario: An unrelated input edit preserves all formula caches
    Given a workbook with input one and a static cross-sheet formula chain
    When I change an unrelated numeric cell with explicit invalidation
    Then cache payloads and workbook calculation metadata remain byte-identical

  @CACHE-003
  Scenario Outline: Unproved dependent structures refuse the complete edit
    Given a formula workbook containing "<condition>"
    When I attempt an explicit invalidating input edit
    Then invalidation returns a typed refusal without any package mutation

    Examples:
      | condition |
      | dynamic formula |
      | shared formula |
      | unknown sheet |
      | existing calculation chain |
      | manual calculation mode |
      | locked workbook |

  @CACHE-004
  Scenario: An empty inactive workbook protection marker permits invalidation
    Given a formula workbook containing "inactive workbook protection"
    When I set the input to ten with explicit invalidation
    Then direct and transitive cached results are absent or empty
    And the workbook requests full recalculation without computing an answer
