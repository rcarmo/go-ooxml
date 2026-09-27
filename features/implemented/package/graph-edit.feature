@implemented @go @package
Feature: Planned part addition and relationship retargeting
  New payloads and registry patches commit together without rewriting shared old payloads.

  @SURGERY-001
  Scenario: Add a private payload and retarget one of two shared edges
    Given a retained package with two edges sharing a payload
    When I plan a new payload for the first edge
    Then planning leaves the original archive byte-identical
    When I apply and deliver the graph plan
    Then only the selected edge points to the new registered payload
    And unrelated registry bytes and original payloads are preserved

  @SURGERY-002
  Scenario Outline: Unproved graph changes refuse the entire plan
    Given a retained package with two edges sharing a payload
    When I plan a graph change with "<condition>"
    Then graph planning refuses without adding parts or changing bytes

    Examples:
      | condition |
      | case colliding part |
      | absent target |
      | external edge |
      | duplicate edge edit |
      | reserved registry addition |
      | malformed XML addition |

  @SURGERY-003
  Scenario: A graph plan cannot overwrite an intervening payload edit
    Given a retained package with two edges sharing a payload
    When I plan a new payload for the first edge
    And I change an existing payload before applying the graph plan
    Then the stale graph plan refuses and preserves the intervening edit
