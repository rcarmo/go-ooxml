@implemented @go @package
Feature: Retargeting preserves the package-relative or absolute target form
  A new target must resolve from its actual source while retaining the input path form.

  @TARGETFORM-001
  Scenario Outline: Retarget paths retain their original form
    Given a graph edge using "<form>" target form
    When I retarget it to media containing escaped characters
    Then its new target has "<form>" form and resolves to the exact added member
    And the old media and all unrelated payloads remain unchanged

    Examples:
      | form |
      | relative |
      | absolute |
