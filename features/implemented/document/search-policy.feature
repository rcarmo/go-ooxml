@implemented @go @document
Feature: Explicit Word search policy and context ranking
  Normalisation selects original text; proximity never hides ambiguous candidates.

  @SEARCH-001
  Scenario: Normalised matching retains original typographic text
    Given a Word paragraph containing "Straße – “Cost”"
    When I search with normalisation for "STRASSE - “cost”"
    Then the search evidence is exactly "Straße – “Cost”"
    And that selected evidence can authorise an ordinary replacement

  @SEARCH-002
  Scenario: Casefold expansion cannot select half an original character
    Given a Word paragraph containing "ß"
    When I search with normalisation for "s"
    Then normalised search returns no partial-character target

  @SEARCH-003
  Scenario: Near ranks every candidate but does not discard distant matches
    Given Word search paragraphs "red far" and "red context"
    When I rank matches for "red" near "context"
    Then both candidates remain and the contextual paragraph ranks first

  @SEARCH-004
  Scenario: Conflicting positional and contextual selectors refuse
    Given a Word paragraph containing "red red"
    When I combine an explicit occurrence with contextual ranking
    Then the search argument conflict is reported
