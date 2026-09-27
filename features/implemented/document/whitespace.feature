@implemented @go @document
Feature: Word whitespace preservation on successful edits
  Changed text receives explicit XML space preservation without altering unrelated attributes.

  @SPACE-001
  Scenario: A single replacement adds the required preservation attribute
    Given a Word text leaf without XML space preservation
    When I replace its text with leading and trailing spaces
    Then the changed leaf has XML space preserve and exact new text
    And its unrelated attribute bytes remain unchanged

  @SPACE-002
  Scenario: A selected batch adds one preservation attribute to a shared leaf
    Given two Word matches in a leaf without XML space preservation
    When a selected batch adds significant spaces to both matches
    Then the saved leaf has one XML space preserve attribute
    And both replacements retain their significant spaces
