@implemented @go @spreadsheet
Feature: Remove an obsolete calculation chain with dependent-cache invalidation
  The optional chain is removed only when a supported numeric edit affects formulas.

  @CHAIN-001
  Scenario: Dependent input edit atomically removes an owned calculation chain
    Given a formula workbook containing "owned calculation chain"
    When I set the input to ten with explicit invalidation
    Then direct and transitive cached results are absent or empty
    And the workbook requests full recalculation without computing an answer
    And the calculation chain part and its registrations are removed
    And unrelated cached results and all formula text remain unchanged

  @CHAIN-002
  Scenario: An unrelated numeric edit retains the chain and its registrations
    Given a formula workbook containing "owned calculation chain"
    When I change an unrelated numeric cell with explicit invalidation
    Then cache payloads and workbook calculation metadata remain byte-identical
    And the calculation chain and registry payloads remain byte-identical

  @CHAIN-003
  Scenario: A same-value edit retains the chain and keeps its target reusable
    Given a formula workbook containing "owned calculation chain"
    When I apply a same-value edit then reuse its numeric target
    Then the calculation chain part and its registrations are removed

  @CHAIN-004
  Scenario Outline: Unknown or shared chain ownership refuses without editing
    Given a formula workbook containing "<condition>"
    When I attempt an explicit invalidating input edit
    Then invalidation returns a typed refusal without any package mutation

    Examples:
      | condition |
      | shared calculation chain |
      | extended calculation chain |
      | external calculation chain |
      | chain with outgoing relationships |
