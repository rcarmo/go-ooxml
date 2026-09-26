@implemented @go @package
Feature: Read-only relationship ownership inspection
  Package inspection resolves local ownership without fetching external targets.

  @GRAPH-001
  Scenario: Shared and external relationships retain distinct ownership
    Given a package with two users of one image and an external hyperlink
    When I inspect the retained package relationship graph
    Then the image has two incoming relationships
    And the hyperlink remains an unresolved external edge
    And graph inspection leaves the archive byte-identical

  @GRAPH-002
  Scenario Outline: Ambiguous or dangling registries refuse inspection
    Given a relationship package with "<defect>"
    When I inspect the retained package relationship graph
    Then graph inspection returns a relationship-policy refusal

    Examples:
      | defect |
      | missing target |
      | duplicate relationship ID |
      | escaping target |
      | duplicate content-type default |
      | unknown registry extension |
