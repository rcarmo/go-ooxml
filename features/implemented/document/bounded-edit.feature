@implemented @go @document
Feature: Bounded preservation-safe Word text correction
  Plain body text corrections preserve every unrelated XML byte and package part.

  @WORD-001
  Scenario: A plain run correction preserves opaque markup and media
    Given a Word package containing a plain body text run and opaque content
    When I replace the unique text "old text" with "new & text" through a guarded target
    And save the edited Word package
    Then only the matched XML character data differs
    And every other package member retains its payload bytes
    And reusing the consumed target returns a stale-target refusal

  @WORD-002
  Scenario Outline: Unsupported ownership or editing policy refuses unchanged
    Given a Word package with "<condition>"
    When I request a guarded correction of "old text"
    Then the correction returns a typed refusal
    And saving the Word session retains the original archive bytes

    Examples:
      | condition |
      | duplicate text |
      | active document protection |
      | tracked revisions |
      | field boundary |
      | unknown settings extension |
