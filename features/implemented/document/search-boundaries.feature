@implemented @go @document
Feature: Word normalised and paragraph-boundary edge cases
  Failed folded candidates cannot hide later whole-character matches.

  @EDGE-001
  Scenario: A partial folded candidate does not consume a later valid character match
    Given a Word paragraph containing "sß"
    When I search with normalisation for "ss"
    Then the search evidence is exactly "ß"

  @EDGE-002
  Scenario: A span ending in a paragraph separator remains inspection-only
    Given two scoped Word paragraphs containing "first" and "second"
    When I search for text ending at the paragraph separator
    Then attempting to edit that cross-paragraph match refuses
