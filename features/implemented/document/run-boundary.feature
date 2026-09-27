@implemented @go @document
Feature: Word insertion at run boundaries
  Insertion needs identical complete direct formatting and a proved shared owner.

  @INSERT-001
  Scenario: Identically formatted adjacent runs allow a boundary insertion
    Given adjacent Word runs with identical direct formatting
    When I insert a character at their shared text boundary
    Then the combined text includes the insertion with unchanged run properties

  @INSERT-002
  Scenario Outline: Differing or interrupted boundary ownership refuses
    Given adjacent Word runs with "<condition>"
    When I attempt an insertion at their shared text boundary
    Then the insertion refuses and retains the original archive

    Examples:
      | condition |
      | different direct formatting |
      | an intervening bookmark |
      | different run attributes |
