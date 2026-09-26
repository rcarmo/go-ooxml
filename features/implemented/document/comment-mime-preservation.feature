@implemented @go @document
Feature: Unrelated edits preserve differing extended-comment content types
  Observed source-library disagreement does not authorise silent content-type normalization.

  @MIME-001
  Scenario Outline: Preserve the input spelling without claiming Office certification
    Given a Word package with extended-comment MIME spelling "<source>"
    When I make an unrelated guarded body text correction
    Then the delivered package retains the exact extended-comment MIME spelling
    And extended comments and relationship payloads remain byte-identical

    Examples:
      | source |
      | pinned-go |
      | specified-metadata |
