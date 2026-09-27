@implemented @go @document
Feature: Word search story and view scope
  Paragraph boundaries are explicit characters and historical evidence cannot authorise current edits.

  @SCOPE-001
  Scenario: Search sees exactly one newline between paragraphs
    Given two scoped Word paragraphs containing "first" and "second"
    When I search for their text with an explicit paragraph separator
    Then one cross-paragraph match reports the exact combined text
    And attempting to edit that cross-paragraph match refuses

  @SCOPE-002
  Scenario: Historical-view matches are read-only
    Given scoped Word text with a tracked deletion
    When I search original-view text for "removed"
    Then the historical match reports "removed"
    And editing the historical match refuses without mutation

  @SCOPE-003
  Scenario: An explicit related story does not search the body instead
    Given scoped Word body and header text both equal to "repeat"
    When I search only the header story for "repeat"
    Then exactly one match identifies the header story
