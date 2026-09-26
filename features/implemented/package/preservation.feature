@implemented @go @package
Feature: Retained-source package editing
  A bounded part replacement preserves every other source payload.

  @SOURCE-001
  Scenario: A no-op save preserves the exact source archive
    Given an Office source archive with an opaque custom part
    When I open a retained-source editing session and save without edits
    Then the output archive is byte-identical to the source

  @SOURCE-002
  Scenario: Replacing one part leaves all other member payloads unchanged
    Given an Office source archive with an opaque custom part
    When I replace the main XML part using its current fingerprint
    Then only that member payload changes after saving
    And the retained source bytes remain unchanged

  @SOURCE-003
  Scenario: A stale replacement batch leaves all parts unchanged
    Given an Office source archive with an opaque custom part
    When I request two replacements with one stale fingerprint
    Then the batch returns a stale-target refusal
    And the session still saves the exact original archive
