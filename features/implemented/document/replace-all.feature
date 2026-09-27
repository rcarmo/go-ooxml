@implemented @go @document
Feature: Word replace-all reports independent refusals
  Supported matches commit together while unsupported matches remain unchanged and reported.

  @ALL-001
  Scenario: A wrapper refusal does not discard independent supported matches
    Given two plain Word matches and one hyperlink-owned match
    When I replace all occurrences of "red" with "blue"
    Then two matches change and one unsupported match is reported
    And the hyperlink-owned text remains unchanged

  @ALL-002
  Scenario: Replace-all skips matches already equal to the replacement
    Given two plain Word matches and one hyperlink-owned match
    When I replace all occurrences of "red" with "red"
    Then all three matches are skipped without refusals or mutation

  @ALL-003
  Scenario: Invalid XML replacement cannot commit any selected match
    Given two plain Word matches and one hyperlink-owned match
    When I replace all occurrences with an XML-illegal value
    Then no match changes and the original package bytes remain exact
